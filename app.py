import os
import json
import time
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, Response, stream_with_context, jsonify, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import anthropic
import openai
from google import generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import stripe
import PyPDF2

# Load environment variables
load_dotenv()

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key-change-me')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URI', 'sqlite:///ai_architect.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Initialize extensions
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Stripe configuration
stripe_key = os.getenv('STRIPE_SECRET_KEY', '')
if stripe_key:
    stripe.api_key = stripe_key
    # Debug print with masked key
    masked_key = stripe_key[:7] + '...' + stripe_key[-4:] if len(stripe_key) > 11 else '***'
    print(f"✅ Stripe Key Loaded: {masked_key}")
else:
    print("⚠️ WARNING: STRIPE_SECRET_KEY not found in .env")

DOMAIN = os.getenv('DOMAIN', 'http://localhost:5000')

# Configuration
MAX_ROUNDS = int(os.getenv('MAX_ROUNDS', 10))
SUCCESS_THRESHOLD = int(os.getenv('SUCCESS_THRESHOLD', 90))
STAGNATION_THRESHOLD = float(os.getenv('STAGNATION_THRESHOLD', 1.5))
HISTORY_PRUNING_START = int(os.getenv('HISTORY_PRUNING_START', 4))
ALLOWED_EXTENSIONS = {'pdf', 'docx'}

# Pricing tiers
PRICING_TIERS = {
    '2000': {'price': 2000, 'display': '£20', 'words': 2000},  # Price in cents
    '4000': {'price': 4000, 'display': '£40', 'words': 4000}
}

# Initialize AI clients
anthropic_client = anthropic.Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))
openai.api_key = os.getenv('OPENAI_API_KEY')
genai.configure(api_key=os.getenv('GOOGLE_API_KEY'))

# Database Models
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    essays = db.relationship('Essay', backref='author', lazy=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class Essay(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    instructions = db.Column(db.Text, nullable=False)
    final_content = db.Column(db.Text, nullable=False)
    final_score = db.Column(db.Integer, nullable=False)
    rounds_used = db.Column(db.Integer, nullable=False)
    stop_reason = db.Column(db.String(100), nullable=False)
    word_count_limit = db.Column(db.Integer, default=2000)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# File handling functions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text.strip()
    except Exception as e:
        raise Exception(f"Error reading PDF: {str(e)}")

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text.strip()
    except Exception as e:
        raise Exception(f"Error reading DOCX: {str(e)}")

def parse_uploaded_file(file_path, filename):
    """Parse uploaded file and extract text"""
    ext = filename.rsplit('.', 1)[1].lower()
    
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    else:
        raise Exception("Unsupported file type")

# AI Helper Functions with Retry Logic
def call_with_retry(func, max_retries=3, delay=2):
    """Generic retry wrapper for API calls"""
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(delay * (attempt + 1))
    return None

def call_claude_writer(prompt, max_tokens=4000):
    """Claude 3.5 Sonnet - Writer"""
    def api_call():
        response = anthropic_client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    return call_with_retry(api_call)

def call_gpt4_examiner(prompt):
    """GPT-4o - Examiner"""
    def api_call():
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content
    return call_with_retry(api_call)

def call_gemini_reviewer(prompt):
    """Gemini 1.5 Pro - Reviewer"""
    def api_call():
        model = genai.GenerativeModel('gemini-1.5-pro')
        response = model.generate_content(prompt)
        return response.text
    return call_with_retry(api_call)

def extract_score(examiner_response):
    """Extract numerical score from examiner response"""
    try:
        import re
        matches = re.findall(r'\b(\d{1,3})\b', examiner_response)
        for match in matches:
            score = int(match)
            if 0 <= score <= 100:
                return score
        return 0
    except:
        return 0

def build_context(round_num, original_instructions, current_draft, previous_feedback, full_history):
    """Build context based on round number (implements pruning from round 4)"""
    if round_num < HISTORY_PRUNING_START:
        return full_history
    else:
        pruned_context = f"""ORIGINAL INSTRUCTIONS:
{original_instructions}

CURRENT DRAFT:
{current_draft}

PREVIOUS ROUND FEEDBACK:
{previous_feedback}
"""
        return pruned_context

# Main AI Loop Generator (SSE)
def generate_essay_stream(instructions, word_count_limit):
    """Main AI loop with streaming logs"""
    yield f"data: {json.dumps({'type': 'log', 'message': f'🚀 Initializing AI Assignment Architect (Target: {word_count_limit} words)...', 'agent': 'system'})}\n\n"
    
    # Step A: Initial Draft by Claude
    yield f"data: {json.dumps({'type': 'log', 'message': '📝 Writer Agent activating...', 'agent': 'writer'})}\n\n"
    
    writer_prompt = f"""You are an expert academic writer. Create a high-quality essay based on these instructions:

{instructions}

IMPORTANT: Target word count is approximately {word_count_limit} words. Write a comprehensive, well-structured essay that fully addresses all requirements. Use proper academic language, clear arguments, and supporting evidence."""

    try:
        current_draft = call_claude_writer(writer_prompt)
        yield f"data: {json.dumps({'type': 'log', 'message': '✅ Writer has completed the initial draft.', 'agent': 'writer'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'❌ Writer error: {str(e)}', 'agent': 'writer'})}\n\n"
        return

    # Initialize tracking variables
    scores = []
    full_history = f"Original Instructions:\n{instructions}\n\nDraft 1:\n{current_draft}\n\n"
    best_draft = current_draft
    best_score = 0
    previous_feedback = ""
    stop_reason = ""

    # Step B: Refinement Loop
    yield f"data: {json.dumps({'type': 'log', 'message': f'🔄 Starting refinement loop (Maximum {MAX_ROUNDS} rounds)...', 'agent': 'system'})}\n\n"
    
    for round_num in range(1, MAX_ROUNDS + 1):
        yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ Round {round_num}/{MAX_ROUNDS} ━━━', 'agent': 'system'})}\n\n"
        
        # 1. Examiner (GPT-4o) grades
        yield f"data: {json.dumps({'type': 'log', 'message': '🎯 Examiner Agent analyzing quality...', 'agent': 'examiner'})}\n\n"
        
        examiner_prompt = f"""You are a strict academic examiner. Grade this essay on a scale of 0-100.

Instructions:
{instructions}

Essay:
{current_draft}

Provide a score (0-100) and brief justification. Format: "Score: [number]" followed by reasoning."""

        try:
            examiner_response = call_gpt4_examiner(examiner_prompt)
            score = extract_score(examiner_response)
            scores.append(score)
            
            yield f"data: {json.dumps({'type': 'score', 'round': round_num, 'score': score, 'agent': 'examiner'})}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Examiner verdict: {score}/100', 'agent': 'examiner'})}\n\n"
            
            if score > best_score:
                best_score = score
                best_draft = current_draft
                
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Examiner error: {str(e)}', 'agent': 'examiner'})}\n\n"
            score = 0
            scores.append(0)

        # Break Condition A: SUCCESS
        if score >= SUCCESS_THRESHOLD:
            stop_reason = f"SUCCESS: Score {score} >= {SUCCESS_THRESHOLD}"
            yield f"data: {json.dumps({'type': 'log', 'message': f'🎉 {stop_reason}! Quality threshold achieved.', 'agent': 'system'})}\n\n"
            break

        # Break Condition B: STAGNATION
        if round_num >= 3:
            recent_scores = scores[-2:]
            if len(recent_scores) == 2:
                improvement = recent_scores[-1] - recent_scores[-2]
                if improvement < STAGNATION_THRESHOLD:
                    stop_reason = f"STAGNATION: Improvement {improvement:.1f} < {STAGNATION_THRESHOLD}"
                    yield f"data: {json.dumps({'type': 'log', 'message': f'⏸️ {stop_reason}. Optimization complete.', 'agent': 'system'})}\n\n"
                    break

        # 2. Reviewer (Gemini) provides critique
        yield f"data: {json.dumps({'type': 'log', 'message': '💬 Reviewer Agent providing feedback...', 'agent': 'reviewer'})}\n\n"
        
        reviewer_prompt = f"""You are an expert writing coach. Review this essay and provide specific, actionable feedback for improvement.

Instructions:
{instructions}

Current Essay:
{current_draft}

Current Score: {score}/100
Target Word Count: {word_count_limit} words

Provide detailed critique focusing on:
1. Content gaps or weaknesses
2. Structure and flow issues
3. Language and style improvements
4. Specific suggestions for the next revision"""

        try:
            reviewer_response = call_gemini_reviewer(reviewer_prompt)
            previous_feedback = reviewer_response
            yield f"data: {json.dumps({'type': 'log', 'message': f'📋 Reviewer feedback received.', 'agent': 'reviewer'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Reviewer error: {str(e)}', 'agent': 'reviewer'})}\n\n"
            previous_feedback = "No feedback available due to error."

        # Break Condition C: LIMIT
        if round_num == MAX_ROUNDS:
            stop_reason = f"LIMIT: Reached maximum {MAX_ROUNDS} rounds"
            yield f"data: {json.dumps({'type': 'log', 'message': f'🛑 {stop_reason}.', 'agent': 'system'})}\n\n"
            break

        # 3. Writer refines based on feedback
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Writer refining based on feedback...', 'agent': 'writer'})}\n\n"
        
        context = build_context(round_num, instructions, current_draft, previous_feedback, full_history)
        
        if round_num >= HISTORY_PRUNING_START:
            yield f"data: {json.dumps({'type': 'log', 'message': f'💾 Context optimization active (Round {round_num}+)', 'agent': 'system'})}\n\n"
        
        refine_prompt = f"""{context}

Based on the feedback above, rewrite and improve the essay. Address all critiques and aim for a score above {SUCCESS_THRESHOLD}. Maintain approximately {word_count_limit} words."""

        try:
            current_draft = call_claude_writer(refine_prompt)
            full_history += f"Round {round_num} Feedback:\n{previous_feedback}\n\nRevised Draft {round_num + 1}:\n{current_draft}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Writer has refined the draft.', 'agent': 'writer'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'❌ Writer error: {str(e)}', 'agent': 'writer'})}\n\n"
            break

    # Final results
    if not stop_reason:
        stop_reason = "COMPLETED: All rounds finished"
    
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 Process complete! {stop_reason}', 'agent': 'system'})}\n\n"
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏆 Best Score: {best_score}/100 | Rounds: {len(scores)}', 'agent': 'system'})}\n\n"
    
    # Save to database
    try:
        essay = Essay(
            user_id=current_user.id,
            title=instructions[:100],
            instructions=instructions,
            final_content=best_draft,
            final_score=best_score,
            rounds_used=len(scores),
            stop_reason=stop_reason,
            word_count_limit=word_count_limit
        )
        db.session.add(essay)
        db.session.commit()
        
        yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': best_score, 'rounds': len(scores)})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'Database error: {str(e)}', 'agent': 'system'})}\n\n"

# Routes
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user)
            flash('Login successful!', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid username or password', 'error')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        
        if User.query.filter_by(username=username).first():
            flash('Username already exists', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email already registered', 'error')
            return redirect(url_for('register'))
        
        user = User(username=username, email=email)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        
        flash('Registration successful! Please login.', 'success')
        return redirect(url_for('login'))
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Logged out successfully', 'success')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    essays = Essay.query.filter_by(user_id=current_user.id).order_by(Essay.created_at.desc()).all()
    return render_template('dashboard.html', essays=essays, user=current_user)

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    """Handle file upload and text extraction"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only PDF and DOCX allowed.'}), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Extract text from file
        extracted_text = parse_uploaded_file(filepath, filename)
        
        # Clean up uploaded file
        os.remove(filepath)
        
        return jsonify({'success': True, 'text': extracted_text})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/create-checkout-session', methods=['POST'])
@login_required
def create_checkout_session():
    """Create Stripe checkout session with dynamic pricing"""
    try:
        data = request.get_json()
        word_count = data.get('word_count', '2000')
        instructions = data.get('instructions', '')
        
        if word_count not in PRICING_TIERS:
            return jsonify({'error': 'Invalid word count tier'}), 400
        
        tier = PRICING_TIERS[word_count]
        
        # Store in session for post-payment processing
        session['pending_essay'] = {
            'instructions': instructions,
            'word_count': tier['words']
        }
        
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price_data': {
                    'currency': 'gbp',
                    'unit_amount': tier['price'],
                    'product_data': {
                        'name': f'AI Essay Generation ({tier["words"]} words)',
                        'description': f'Professional AI-powered essay up to {tier["words"]} words',
                    },
                },
                'quantity': 1,
            }],
            mode='payment',
            allow_promotion_codes=True,
            success_url=DOMAIN + '/payment-success?session_id={CHECKOUT_SESSION_ID}',
            cancel_url=DOMAIN + '/dashboard',
            client_reference_id=str(current_user.id),
            customer_email=current_user.email,
        )
        
        return jsonify({'id': checkout_session.id})
    except Exception as e:
        print(f"❌ Stripe Error: {str(e)}")
        return jsonify({'error': str(e)}), 403

@app.route('/payment-success')
@login_required
def payment_success():
    """Handle successful payment and start generation"""
    session_id = request.args.get('session_id')
    
    if session_id:
        try:
            checkout_session = stripe.checkout.Session.retrieve(session_id)
            
            if checkout_session.payment_status == 'paid':
                flash('Payment successful! Starting essay generation...', 'success')
                # Redirect to generation page
                return redirect(url_for('dashboard') + '?start_generation=true')
            else:
                flash('Payment verification failed.', 'error')
        except Exception as e:
            flash(f'Error verifying payment: {str(e)}', 'error')
    
    return redirect(url_for('dashboard'))

@app.route('/generate', methods=['POST'])
@login_required
def generate():
    """Generate essay after payment"""
    instructions = request.form.get('instructions')
    word_count = int(request.form.get('word_count', 2000))
    
    if not instructions:
        return jsonify({'error': 'Instructions required'}), 400
    
    return Response(
        stream_with_context(generate_essay_stream(instructions, word_count)),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no'
        }
    )

@app.route('/download/<int:essay_id>')
@login_required
def download(essay_id):
    essay = Essay.query.get_or_404(essay_id)
    
    if essay.user_id != current_user.id:
        flash('Unauthorized access', 'error')
        return redirect(url_for('dashboard'))
    
    # Create Word document
    doc = Document()
    
    # Title
    title = doc.add_heading(essay.title, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Metadata
    doc.add_paragraph(f"Generated: {essay.created_at.strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Final Score: {essay.final_score}/100")
    doc.add_paragraph(f"Rounds Used: {essay.rounds_used}")
    doc.add_paragraph(f"Word Count Limit: {essay.word_count_limit}")
    doc.add_paragraph(f"Stop Reason: {essay.stop_reason}")
    doc.add_paragraph("")
    
    # Instructions
    doc.add_heading('Original Instructions', 1)
    doc.add_paragraph(essay.instructions)
    doc.add_paragraph("")
    
    # Content
    doc.add_heading('Final Essay', 1)
    doc.add_paragraph(essay.final_content)
    
    # Save to BytesIO
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return Response(
        file_stream.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers={
            'Content-Disposition': f'attachment; filename=essay_{essay_id}.docx'
        }
    )

@app.route('/view/<int:essay_id>')
@login_required
def view_essay(essay_id):
    essay = Essay.query.get_or_404(essay_id)
    
    if essay.user_id != current_user.id:
        flash('Unauthorized access', 'error')
        return redirect(url_for('dashboard'))
    
    return jsonify({
        'title': essay.title,
        'instructions': essay.instructions,
        'content': essay.final_content,
        'score': essay.final_score,
        'rounds': essay.rounds_used,
        'word_count_limit': essay.word_count_limit,
        'stop_reason': essay.stop_reason,
        'created_at': essay.created_at.strftime('%Y-%m-%d %H:%M')
    })

# Initialize database
with app.app_context():
    db.create_all()
    
    # Create demo user (no free credits)
    if not User.query.filter_by(username='demo').first():
        demo_user = User(username='demo', email='demo@example.com')
        demo_user.set_password('demo123')
        db.session.add(demo_user)
        db.session.commit()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
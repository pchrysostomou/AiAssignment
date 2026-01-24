import os
import json
import time
import re
import requests
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, Response, stream_with_context, jsonify, session, send_file
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
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.units import inch

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key-change-me')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URI', 'sqlite:///academic_suite.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['REPORTS_FOLDER'] = 'reports'
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25MB

# Add Stripe Publishable Key to config so templates can access it
app.config['STRIPE_PUBLISHABLE_KEY'] = os.getenv('STRIPE_PUBLISHABLE_KEY', '')

# CRITICAL FIX: Set permissive CSP that allows Stripe.js and eval() for payment processing
@app.after_request
def add_security_headers(response):
    """
    Configure Content Security Policy to allow:
    1. Stripe.js library and checkout iframe
    2. unsafe-eval for Stripe's payment processing
    3. Our own domain for API calls
    """
    # Get the current domain
    domain = os.getenv('DOMAIN', 'https://aiassignment-x631.onrender.com')
    
    # Build CSP policy that allows Stripe while maintaining reasonable security
    csp_policy = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' 'unsafe-eval' https://js.stripe.com https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data: https:; "
        f"connect-src 'self' {domain} https://api.stripe.com https://www.googleapis.com; "
        "frame-src 'self' https://js.stripe.com https://hooks.stripe.com; "
        "frame-ancestors 'self';"
    )
    
    # Set CSP header
    response.headers['Content-Security-Policy'] = csp_policy
    
    # Add CORS headers for API routes
    if request.path.startswith('/api/') or request.path.startswith('/generate-') or request.path.startswith('/create-checkout') or request.path.startswith('/check-'):
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
        response.headers['Access-Control-Max-Age'] = '3600'
    
    # Prevent caching for API routes
    if request.path.startswith('/api/') or request.path.startswith('/generate-') or request.path.startswith('/create-checkout'):
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
    
    return response

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORTS_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Stripe initialization - uses environment variable
stripe_key = os.getenv('STRIPE_SECRET_KEY', '')
if stripe_key:
    stripe.api_key = stripe_key
    print(f"✅ Stripe Secret Key Loaded (ending: ...{stripe_key[-4:]})")
else:
    print("⚠️ WARNING: STRIPE_SECRET_KEY not found in environment variables")

# Log Stripe Publishable Key status
stripe_pub_key = os.getenv('STRIPE_PUBLISHABLE_KEY', '')
if stripe_pub_key:
    print(f"✅ Stripe Publishable Key Loaded (ending: ...{stripe_pub_key[-4:]})")
else:
    print("⚠️ WARNING: STRIPE_PUBLISHABLE_KEY not found in environment variables")

# Hardcode production domain for Stripe redirects
DOMAIN = os.getenv('DOMAIN', 'https://aiassignment-x631.onrender.com')

# Updated pricing structure
PRICING = {
    'writer': {'price_per_1k': 1000, 'display': '£10 per 1,000 words'},  # Dynamic pricing
    'plagiarism': {'price': 1000, 'display': '£10'},
    'ai_check': {'price': 1000, 'display': '£10'},
    'grader': {'price': 1000, 'display': '£10'}  # Updated to £10
}

# API Credentials Initialization with Validation
anthropic_api_key = os.getenv('ANTHROPIC_API_KEY', '')
if anthropic_api_key:
    # Initialize with STRICT API version header
    anthropic_client = anthropic.Anthropic(
        api_key=anthropic_api_key,
        default_headers={"anthropic-version": "2023-06-01"}
    )
    print(f"✅ Anthropic API Key Loaded (ending: ...{anthropic_api_key[-4:]})")
    print(f"✅ Anthropic Client Initialized with API version: 2023-06-01")
else:
    anthropic_client = None
    print("⚠️ WARNING: ANTHROPIC_API_KEY not found in environment variables")

openai_api_key = os.getenv('OPENAI_API_KEY', '')
if openai_api_key:
    openai.api_key = openai_api_key
    print(f"✅ OpenAI API Key Loaded (ending: ...{openai_api_key[-4:]})")
else:
    print("⚠️ WARNING: OPENAI_API_KEY not found in environment variables")

google_api_key = os.getenv('GOOGLE_API_KEY', '')
if google_api_key:
    genai.configure(api_key=google_api_key)
    print(f"✅ Google API Key Loaded (ending: ...{google_api_key[-4:]})")
else:
    print("⚠️ WARNING: GOOGLE_API_KEY not found in environment variables")

# Google Custom Search Engine Configuration - Force string type
GOOGLE_CSE_ID = str(os.getenv('GOOGLE_CSE_ID', '50f552d88c3e14772')).strip()
if GOOGLE_CSE_ID and GOOGLE_CSE_ID != '':
    print(f"✅ Google CSE ID Configured: '{GOOGLE_CSE_ID}' (type: {type(GOOGLE_CSE_ID).__name__}, length: {len(GOOGLE_CSE_ID)})")
else:
    print("⚠️ WARNING: GOOGLE_CSE_ID not found or empty in environment variables")

# MOCK EMAIL FUNCTION (For localhost development)
def send_email_mock(user_email, subject, body):
    """Mock email sender - prints to console instead of sending real emails"""
    print("\n" + "="*80)
    print("📧 MOCK EMAIL SENT (Console Log)")
    print("="*80)
    print(f"To: {user_email}")
    print(f"Subject: {subject}")
    print(f"Body:\n{body}")
    print("="*80 + "\n")

# Database Models
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    essays = db.relationship('Essay', backref='author', lazy=True)
    reports = db.relationship('Report', backref='author', lazy=True)
    payments = db.relationship('Payment', backref='user', lazy=True)

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
    word_count_limit = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    tool_type = db.Column(db.String(50), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    result_data = db.Column(db.Text, nullable=False)
    file_path = db.Column(db.String(500))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    amount = db.Column(db.Integer, nullable=False)
    tool_type = db.Column(db.String(50), nullable=False)
    stripe_session_id = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# File handling
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'pdf', 'docx', 'txt'}

def extract_text_from_pdf(file_path):
    """
    Extract text from PDF file with MEMORY-EFFICIENT streaming approach.
    Limits pages to prevent RAM overload on limited server.
    """
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            
            # MEMORY OPTIMIZATION: Limit to 100 pages for 25MB files
            max_pages = min(total_pages, 100)
            
            # Extract text page by page (memory efficient)
            text_chunks = []
            for i in range(max_pages):
                try:
                    page_text = pdf_reader.pages[i].extract_text()
                    if page_text:
                        text_chunks.append(page_text)
                except Exception as e:
                    print(f"⚠️ Warning: Could not extract page {i+1}: {str(e)}")
                    continue
            
            text = "\n".join(text_chunks)
            
            if total_pages > max_pages:
                text += f"\n\n[Note: Document has {total_pages} pages. First {max_pages} pages extracted for performance on limited RAM server.]"
            
            return text.strip()
    except Exception as e:
        raise Exception(f"PDF extraction error: {str(e)}")

def extract_text_from_docx(file_path):
    """
    Extract text from DOCX file with MEMORY-EFFICIENT approach.
    Limits paragraphs to prevent RAM overload.
    """
    try:
        doc = Document(file_path)
        total_paragraphs = len(doc.paragraphs)
        
        # MEMORY OPTIMIZATION: Limit to 300 paragraphs for 25MB files
        max_paragraphs = min(total_paragraphs, 300)
        
        # Extract paragraph by paragraph (memory efficient)
        text_chunks = []
        for i in range(max_paragraphs):
            para_text = doc.paragraphs[i].text
            if para_text.strip():
                text_chunks.append(para_text)
        
        text = "\n".join(text_chunks)
        
        if total_paragraphs > max_paragraphs:
            text += f"\n\n[Note: Document has {total_paragraphs} paragraphs. First {max_paragraphs} extracted for performance on limited RAM server.]"
        
        return text.strip()
    except Exception as e:
        raise Exception(f"DOCX extraction error: {str(e)}")

def extract_text_from_file(file_path, filename):
    """
    Helper function to extract text from PDF or DOCX files.
    MEMORY-EFFICIENT: Enforces size limits to prevent RAM overload.
    """
    # Check file size (max 25MB per file)
    file_size = os.path.getsize(file_path)
    file_size_mb = file_size / (1024 * 1024)
    
    if file_size > 25 * 1024 * 1024:  # 25MB
        raise Exception(f"File too large ({file_size_mb:.1f}MB). Maximum 25MB per file.")
    
    print(f"📄 Processing file: {filename} ({file_size_mb:.1f}MB)")
    
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    elif ext == 'txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            # Limit text files to 200KB to prevent RAM issues
            text = f.read(200 * 1024)
            return text.strip()
    raise Exception("Unsupported file type")

def parse_uploaded_file(file_path, filename):
    """Parse uploaded file and extract text"""
    return extract_text_from_file(file_path, filename)

# Google Custom Search Integration - ULTRA FLEXIBLE QUERY BROADENING
def google_search(query, num_results=5):
    """
    Perform Google Custom Search using the configured API key and CSE ID.
    Returns a list of search results with titles, links, and snippets.
    ULTRA FLEXIBLE: Maximum broadening to avoid 'No search results' issue.
    """
    # Verify both API key and CSE ID are valid strings
    if not google_api_key or not GOOGLE_CSE_ID or GOOGLE_CSE_ID == '':
        print(f"⚠️ Google Search unavailable: API key exists: {bool(google_api_key)}, CSE ID: '{GOOGLE_CSE_ID}' (type: {type(GOOGLE_CSE_ID).__name__})")
        return []
    
    try:
        # ULTRA FLEXIBLE: Extract only the most essential keywords
        words = query.split()
        # Take first 3-5 meaningful words (VERY broad, minimal filtering)
        key_words = [w for w in words if len(w) > 1][:5]  # Changed from >2 to >1, reduced from 8 to 5
        
        # If too few keywords, use first 80 chars of original query
        if len(key_words) < 2:
            search_query = query[:80]
        else:
            search_query = ' '.join(key_words)
        
        url = "https://www.googleapis.com/customsearch/v1"
        params = {
            'key': str(google_api_key).strip(),
            'cx': str(GOOGLE_CSE_ID).strip(),
            'q': search_query,
            'num': num_results
        }
        
        print(f"🔍 Google Search Query (ULTRA FLEXIBLE): '{search_query}'")
        print(f"🔍 Using CSE ID: '{GOOGLE_CSE_ID}' (length: {len(GOOGLE_CSE_ID)})")
        
        response = requests.get(url, params=params, timeout=10)
        
        print(f"🔍 Google API Response Status: {response.status_code}")
        
        if response.status_code != 200:
            print(f"❌ Google Search API Error: {response.status_code} - {response.text[:200]}")
            return []
        
        response.raise_for_status()
        
        data = response.json()
        results = []
        
        for item in data.get('items', []):
            results.append({
                'title': item.get('title', ''),
                'link': item.get('link', ''),
                'snippet': item.get('snippet', '')
            })
        
        print(f"✅ Google Search completed: {len(results)} results for '{search_query}'")
        return results
        
    except Exception as e:
        print(f"❌ Google Search error: {str(e)}")
        return []

# AI Helper Functions
def call_with_retry(func, max_retries=3):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(2 * (attempt + 1))

def call_claude(prompt, max_tokens=8000):
    """
    Call Claude Sonnet 3.5 (Tier 1 Primary Model) with retry logic and Haiku fallback.
    PRIMARY: claude-3-5-sonnet-20241022 (High-quality Tier 1 model)
    FALLBACK: claude-3-haiku-20240307 (only after 3 failed Sonnet attempts)
    """
    if not anthropic_client:
        raise Exception("Anthropic API key not configured")
    
    # PRIMARY MODEL: Claude 3.5 Sonnet (Tier 1 High-Quality) with retry logic
    for attempt in range(3):
        try:
            print(f"🤖 Calling Claude Sonnet 3.5 (Tier 1 - attempt {attempt + 1}/3, max_tokens: {max_tokens})")
            response = anthropic_client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}]
            )
            print(f"✅ Claude Sonnet 3.5 API call successful")
            return response.content[0].text
        except anthropic.NotFoundError as e:
            print(f"❌ Claude Sonnet 404 Error (attempt {attempt + 1}/3): {str(e)}")
            if attempt < 2:
                print(f"⏳ Waiting 5 seconds before retry...")
                time.sleep(5)
            else:
                print(f"⚠️ Sonnet failed 3 times, falling back to Haiku...")
        except Exception as e:
            print(f"❌ Claude Sonnet Error (attempt {attempt + 1}/3): {str(e)}")
            if attempt < 2:
                print(f"⏳ Waiting 5 seconds before retry...")
                time.sleep(5)
            else:
                print(f"⚠️ Sonnet failed 3 times, falling back to Haiku...")
    
    # FALLBACK MODEL: Claude 3 Haiku (only if Sonnet fails 3 times)
    try:
        print(f"🤖 FALLBACK: Calling Claude 3 Haiku (max_tokens: 4096)")
        response = anthropic_client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=4096,  # Haiku max is 4096
            messages=[{"role": "user", "content": prompt}]
        )
        print(f"✅ Claude Haiku API call successful (fallback)")
        return response.content[0].text
    except Exception as e:
        print(f"❌ Claude Haiku Error: {str(e)}")
        raise Exception(f"Both Sonnet and Haiku failed: {str(e)}")

def call_gpt4(prompt, model="gpt-4o"):
    """
    Call OpenAI GPT-4o with retry logic and error handling.
    """
    if not openai_api_key:
        raise Exception("OpenAI API key not configured")
    
    def api_call():
        print(f"🤖 Calling OpenAI {model}...")
        response = openai.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        print(f"✅ OpenAI {model} API call successful")
        return response.choices[0].message.content
    return call_with_retry(api_call)

def call_gemini(prompt):
    """
    Call Google Gemini with retry logic and error handling.
    """
    if not google_api_key:
        raise Exception("Google API key not configured")
    
    def api_call():
        print(f"🤖 Calling Google Gemini 1.5 Pro...")
        model = genai.GenerativeModel('gemini-1.5-pro')
        response = model.generate_content(prompt)
        print(f"✅ Google Gemini API call successful")
        return response.text
    return call_with_retry(api_call)

def extract_score(text):
    matches = re.findall(r'\b(\d{1,3})\b', text)
    for match in matches:
        score = int(match)
        if 0 <= score <= 100:
            return score
    return 0

def count_words(text):
    """Count words in text, excluding references section"""
    # Try to find references section
    ref_markers = ['references', 'bibliography', 'works cited']
    text_lower = text.lower()
    
    for marker in ref_markers:
        if f'\n{marker}\n' in text_lower or f'\n{marker}:' in text_lower:
            # Split at references section
            parts = re.split(f'\n{marker}[:\n]', text_lower, maxsplit=1)
            if len(parts) > 1:
                # Count only body text
                body_text = parts[0]
                words = body_text.split()
                return len(words)
    
    # If no references section found, count all words
    words = text.split()
    return len(words)

def check_url_status(url):
    """Check if URL is live (HTTP 200) or broken"""
    try:
        response = requests.head(url, timeout=5, allow_redirects=True)
        return response.status_code == 200
    except:
        return False

# TOOL A: The Detective - 3-Agent Plagiarism Team
def plagiarism_hunter_gemini(text):
    """Step 1: Gemini hunts for sources with web search"""
    prompt = f"""You are The Hunter, a plagiarism detection expert with web search capabilities.

TEXT TO ANALYZE:
{text}

TASKS:
1. Search the web for potential sources of this text
2. Find matching passages
3. Extract URLs of sources
4. For each match, provide:
   - The suspicious passage
   - The source URL
   - Similarity percentage (0-100)

Return JSON format:
{{
    "matches": [
        {{
            "passage": "suspicious text",
            "source_url": "https://...",
            "similarity": 85
        }}
    ]
}}"""
    
    try:
        result = call_gemini(prompt)
        return result
    except Exception as e:
        return json.dumps({"matches": [], "error": str(e)})

def plagiarism_analyst_gpt(student_text, hunter_findings):
    """Step 2: GPT-4o analyzes similarity and paraphrasing"""
    prompt = f"""You are The Analyst, an expert in detecting plagiarism and paraphrasing.

STUDENT TEXT:
{student_text}

SOURCES FOUND BY HUNTER:
{hunter_findings}

TASKS:
1. Compare student text with each source
2. Detect direct copying vs paraphrasing
3. Assign "Suspicion Score" (0-100) for each match
4. Identify specific sentences that are problematic

Return JSON:
{{
    "overall_suspicion": 0-100,
    "analysis": [
        {{
            "student_sentence": "...",
            "source_sentence": "...",
            "suspicion_score": 0-100,
            "type": "DIRECT_COPY/PARAPHRASED/SIMILAR"
        }}
    ]
}}"""
    
    try:
        result = call_gpt4(prompt)
        return result
    except Exception as e:
        return json.dumps({"overall_suspicion": 0, "analysis": [], "error": str(e)})

def plagiarism_reporter_claude(student_text, hunter_data, analyst_data, user_id):
    """Step 3: Claude generates PDF report with highlighted text"""
    filename = f"plagiarism_report_{user_id}_{int(time.time())}.pdf"
    filepath = os.path.join(app.config['REPORTS_FOLDER'], filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], 
                                 fontSize=24, textColor=colors.HexColor('#b537f2'))
    story.append(Paragraph("🔍 The Detective - Plagiarism Report", title_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Metadata
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.2*inch))
    
    # Parse findings
    try:
        hunter_json = json.loads(hunter_data)
        analyst_json = json.loads(analyst_data)
        
        overall_score = analyst_json.get('overall_suspicion', 0)
        verdict_color = colors.red if overall_score > 30 else colors.green
        
        verdict_style = ParagraphStyle('Verdict', parent=styles['Heading2'], textColor=verdict_color)
        story.append(Paragraph(f"Overall Suspicion: {overall_score}%", verdict_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Suspicious passages
        story.append(Paragraph("<b>Suspicious Passages:</b>", styles['Heading3']))
        story.append(Spacer(1, 0.1*inch))
        
        for i, analysis in enumerate(analyst_json.get('analysis', [])[:10], 1):
            student_sent = analysis.get('student_sentence', 'N/A')
            suspicion = analysis.get('suspicion_score', 0)
            match_type = analysis.get('type', 'UNKNOWN')
            
            # Highlight in yellow for high suspicion
            if suspicion > 50:
                highlight_style = ParagraphStyle('Highlight', parent=styles['Normal'],
                                                backColor=colors.yellow)
                story.append(Paragraph(f"<b>Match {i} ({suspicion}% - {match_type}):</b>", styles['Normal']))
                story.append(Paragraph(f'"{student_sent}"', highlight_style))
            else:
                story.append(Paragraph(f"<b>Match {i} ({suspicion}% - {match_type}):</b>", styles['Normal']))
                story.append(Paragraph(f'"{student_sent}"', styles['Italic']))
            
            story.append(Spacer(1, 0.15*inch))
        
        # Reference Status Table
        story.append(PageBreak())
        story.append(Paragraph("<b>Reference Status:</b>", styles['Heading3']))
        story.append(Spacer(1, 0.1*inch))
        
        table_data = [['Source URL', 'Status']]
        for match in hunter_json.get('matches', [])[:15]:
            url = match.get('source_url', 'N/A')
            is_live = check_url_status(url)
            status = '✓ Live' if is_live else '✗ Broken'
            table_data.append([url[:60] + '...' if len(url) > 60 else url, status])
        
        table = Table(table_data, colWidths=[4*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
        
    except Exception as e:
        story.append(Paragraph(f"Error parsing results: {str(e)}", styles['Normal']))
    
    doc.build(story)
    return filename

# TOOL B: The Architect - QUAD AGENT CONSENSUS WRITING SYSTEM (4 AGENTS)
def generate_essay_stream(instructions, word_count):
    yield f"data: {json.dumps({'type': 'log', 'message': f'🎓 The Academic Board convening - QUAD AGENT CONSENSUS MODE ({word_count} words)...'})}\n\n"
    
    # Target 2200 words total to ensure 2000+ body words after references
    target_total_words = 2200
    
    # NEW: Research Phase - Use Google Custom Search for topic research
    yield f"data: {json.dumps({'type': 'log', 'message': '🔍 Researching topic via Google Custom Search...'})}\n\n"
    
    research_context = ""
    try:
        # Extract key topics from instructions for targeted search
        search_query = instructions[:150]
        search_results = google_search(search_query, num_results=5)
        
        if search_results:
            research_context = "\n\nRESEARCH CONTEXT (from web search):\n"
            for idx, result in enumerate(search_results, 1):
                research_context += f"{idx}. {result['title']}\n   Source: {result['link']}\n   Summary: {result['snippet']}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Found {len(search_results)} relevant sources'})}\n\n"
        else:
            yield f"data: {json.dumps({'type': 'log', 'message': '⚠️ No search results, proceeding with general knowledge'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Research phase error: {str(e)}'})}\n\n"
    
    # Initial draft with research context
    yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill drafting initial version...'})}\n\n"
    
    writer_prompt = f"""You are Prof. Quill, operating under the SPARTAN ACADEMIC protocol.

CORE MANDATE: Write with Spartan precision. Be direct, authoritative, and concise. Avoid flowery language. Every sentence must add value.

CRITICAL WORD COUNT RULES:
1. TARGET: {target_total_words} words TOTAL (to ensure {word_count}+ body words after references)
2. BODY TEXT: Aim for {word_count}+ words in Introduction, Analysis, Conclusion
3. EXCLUSION: The References/Bibliography section is EXCLUDED from body word count
4. Write the full essay body first, THEN add a complete References section separately

SMART CITATION LOGIC - Analyze the assignment type and decide:
- CRITICAL REVIEW of a single paper/book: NO external citations needed (only cite the work being reviewed)
- LITERATURE REVIEW: MANY citations needed (10-20+ sources)
- RESEARCH ESSAY: MODERATE citations (5-10 sources)
- OPINION PIECE: FEW citations (2-5 sources for key claims only)
- COMPARATIVE ANALYSIS: MODERATE citations (cite each work being compared)

Based on the instructions below, determine the appropriate citation level and include REAL academic citations ONLY when necessary.

THE BAN LIST - STRICTLY FORBIDDEN WORDS (Essay FAILS if used):
❌ delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust

SPARTAN STYLE RULES:
- Direct, authoritative tone. No hedging ("perhaps", "might", "could be").
- Vary sentence length strategically. Short sentences drive points home. Longer sentences develop complex ideas.
- Avoid excessive transition words. Do NOT overuse: "Furthermore", "Moreover", "In conclusion", "Additionally".
- Use active voice. "The study proves X" not "X is proven by the study".
- Cut unnecessary words. "The fact that" → "That". "In order to" → "To".
- Contractions are acceptable when they strengthen voice (it's, don't, can't).
- Minor stylistic imperfections for authenticity (humans aren't perfect).

INSTRUCTIONS:
{instructions}

{research_context}

Remember: Write {target_total_words} words total to ensure {word_count}+ body words. Apply smart citation logic based on assignment type. No banned words. No fluff."""

    try:
        current_draft = call_claude(writer_prompt, max_tokens=8000)
        current_word_count = count_words(current_draft)
        yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Initial draft complete ({current_word_count} body words)'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'Error: {str(e)}'})}\n\n"
        return

    # QUAD AGENT CONSENSUS LOOP - Minimum 5 rounds, ALL 4 AGENTS MUST SCORE 80+
    MIN_ROUNDS = 5
    MAX_ROUNDS = 15
    
    all_scores = {'quill': [], 'strict': [], 'logic': [], 'chancellor': []}
    best_draft = current_draft
    best_avg_score = 0
    
    for round_num in range(1, MAX_ROUNDS + 1):
        current_word_count = count_words(current_draft)
        yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ ROUND {round_num}/{MAX_ROUNDS} ({current_word_count}/{word_count} body words) ━━━'})}\n\n"
        
        # AGENT 1: Prof. Quill (Claude) - Content & Structure
        yield f"data: {json.dumps({'type': 'log', 'message': '👨‍🏫 Prof. Quill evaluating content & structure...'})}\n\n"
        
        quill_prompt = f"""You are Prof. Quill, an expert academic evaluator. Grade this essay strictly.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Content quality, depth, and originality
- Argument structure and logical flow
- Academic rigor and critical analysis
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected (delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust)
- DEDUCT 15 POINTS if body word count is below {word_count}

Provide a score (0-100) in format: 'Score: [number]' followed by brief feedback."""

        try:
            quill_response = call_claude(quill_prompt, max_tokens=2000)
            quill_score = extract_score(quill_response)
            all_scores['quill'].append(quill_score)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Prof. Quill: {quill_score}/100'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Prof. Quill error: {str(e)}'})}\n\n"
            quill_score = 0
            quill_response = "Error during grading"
            all_scores['quill'].append(0)
        
        # AGENT 2: Dr. Strict (GPT-4o) - Technical Quality
        yield f"data: {json.dumps({'type': 'log', 'message': '👨‍⚖️ Dr. Strict evaluating technical quality...'})}\n\n"
        
        strict_prompt = f"""You are Dr. Strict, a harsh academic grader. Grade this essay strictly.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Grammar, spelling, and punctuation
- Citation format and academic style
- Writing clarity and precision
- Spartan style adherence (direct, concise, no fluff)
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected
- DEDUCT 15 POINTS if body word count is below {word_count}

Provide a score (0-100) in format: 'Score: [number]' followed by brief feedback."""

        try:
            strict_response = call_gpt4(strict_prompt)
            strict_score = extract_score(strict_response)
            all_scores['strict'].append(strict_score)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Dr. Strict: {strict_score}/100'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Dr. Strict error: {str(e)}'})}\n\n"
            strict_score = 0
            strict_response = "Error during grading"
            all_scores['strict'].append(0)
        
        # AGENT 3: Dean Logic (Gemini) - Overall Assessment & Citations
        yield f"data: {json.dumps({'type': 'log', 'message': '👨‍💼 Dean Logic evaluating overall quality & citations...'})}\n\n"
        
        logic_prompt = f"""You are Dean Logic, the final academic authority. Grade this essay strictly.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Overall coherence and persuasiveness
- Citation appropriateness (smart citation logic applied)
- Meeting assignment requirements
- Academic integrity and originality
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected
- DEDUCT 15 POINTS if body word count is below {word_count}

Verify citations based on assignment type:
- Critical review → Minimal citations needed
- Literature review → Many citations needed
- Research essay → Moderate citations needed
- Opinion piece → Few citations needed

Provide a score (0-100) in format: 'Score: [number]' followed by brief feedback."""

        try:
            logic_response = call_gemini(logic_prompt)
            logic_score = extract_score(logic_response)
            all_scores['logic'].append(logic_score)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Dean Logic: {logic_score}/100'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Dean Logic error: {str(e)}'})}\n\n"
            logic_score = 0
            logic_response = "Error during grading"
            all_scores['logic'].append(0)
        
        # AGENT 4: Chancellor GPT (OpenAI GPT-4o) - Final Auditor
        yield f"data: {json.dumps({'type': 'log', 'message': '🎓 Chancellor GPT auditing final quality...'})}\n\n"
        
        chancellor_prompt = f"""You are Chancellor GPT, the supreme academic auditor. Grade this essay with the highest standards.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Overall excellence and scholarly merit
- Argumentation strength and evidence quality
- Professional writing standards
- Readability and engagement
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected
- DEDUCT 15 POINTS if body word count is below {word_count}

As the final auditor, ensure this essay meets the highest academic standards. Be thorough but fair.

Provide a score (0-100) in format: 'Score: [number]' followed by brief feedback."""

        try:
            chancellor_response = call_gpt4(chancellor_prompt)
            chancellor_score = extract_score(chancellor_response)
            all_scores['chancellor'].append(chancellor_score)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Chancellor GPT: {chancellor_score}/100'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Chancellor GPT error: {str(e)}'})}\n\n"
            chancellor_score = 0
            chancellor_response = "Error during grading"
            all_scores['chancellor'].append(0)
        
        # Calculate average score for this round (ALL 4 AGENTS)
        avg_score = round((quill_score + strict_score + logic_score + chancellor_score) / 4)
        yield f"data: {json.dumps({'type': 'score', 'round': round_num, 'score': avg_score})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Round {round_num} Average: {avg_score}/100 (Quill: {quill_score}, Strict: {strict_score}, Logic: {logic_score}, Chancellor: {chancellor_score})'})}\n\n"
        
        # Track best draft
        if avg_score > best_avg_score:
            best_avg_score = avg_score
            best_draft = current_draft
        
        # STRICT TERMINATION CRITERIA - ALL 4 AGENTS MUST SCORE 80+
        # 1. Minimum 5 rounds completed
        # 2. Word count >= 2000 body words
        # 3. ALL FOUR agents score 80+ in the SAME round
        if round_num >= MIN_ROUNDS and current_word_count >= word_count and quill_score >= 80 and strict_score >= 80 and logic_score >= 80 and chancellor_score >= 80:
            yield f"data: {json.dumps({'type': 'log', 'message': f'🎉 CONSENSUS ACHIEVED! All 4 agents agree (80+) after {round_num} rounds!'})}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Final: {avg_score}/100 avg, {current_word_count} body words'})}\n\n"
            break
        
        # Stop if max rounds reached
        if round_num >= MAX_ROUNDS:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⏱️ Max {MAX_ROUNDS} rounds reached. Using best draft (avg: {best_avg_score}/100)'})}\n\n"
            current_draft = best_draft
            break
        
        # REVISION PHASE - All 4 agents contribute feedback
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill revising based on quad-agent feedback...'})}\n\n"
        
        word_count_guidance = ""
        if current_word_count < word_count:
            word_count_guidance = f"\n\nCRITICAL: Current body word count ({current_word_count}) is below target ({word_count}). EXPAND the essay by adding more depth, examples, and analysis. DO NOT just add filler - add substantive content. Aim for {target_total_words} total words."
        elif current_word_count > word_count * 1.15:
            word_count_guidance = f"\n\nNote: Current body word count ({current_word_count}) exceeds target ({word_count}). Tighten the essay by removing redundancy while keeping all key points."
        
        refine_prompt = f"""Revise the essay based on QUAD-AGENT FEEDBACK (4 agents). Apply SPARTAN ACADEMIC protocol.

CRITICAL WORD COUNT RULES (MUST FOLLOW):
1. TARGET: {target_total_words} words TOTAL (to ensure {word_count}+ body words)
2. BODY TEXT: Aim for {word_count}+ words in Introduction, Analysis, Conclusion
3. References/Bibliography is EXCLUDED from body word count
{word_count_guidance}

SMART CITATION LOGIC:
- Analyze assignment type and apply appropriate citation level
- Critical review of single work → Minimal external citations
- Literature review → Many citations
- Research essay → Moderate citations
- Opinion piece → Few citations for key claims only

THE BAN LIST - STRICTLY FORBIDDEN (Essay FAILS if used):
❌ delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust

SPARTAN STYLE:
- Direct, authoritative. No hedging.
- Vary sentence length. Short sentences punch. Longer ones develop.
- Cut transition word overuse (Furthermore, Moreover, Additionally).
- Active voice. Cut fluff.

ORIGINAL INSTRUCTIONS: {instructions}

CURRENT DRAFT: {current_draft}

PROF. QUILL FEEDBACK (Score: {quill_score}/100):
{quill_response}

DR. STRICT FEEDBACK (Score: {strict_score}/100):
{strict_response}

DEAN LOGIC FEEDBACK (Score: {logic_score}/100):
{logic_response}

CHANCELLOR GPT FEEDBACK (Score: {chancellor_score}/100):
{chancellor_response}

Improve the essay while hitting {target_total_words} total words (ensuring {word_count}+ body words after references). Address ALL FOUR agents' feedback. Apply smart citations. Zero banned words. Zero fluff."""

        try:
            current_draft = call_claude(refine_prompt, max_tokens=8000)
            new_word_count = count_words(current_draft)
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Revision complete ({new_word_count} body words)'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': f'Revision error: {str(e)}'})}\n\n"
            break

    final_word_count = count_words(best_draft)
    final_avg = best_avg_score
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 COMPLETE! Final: {final_avg}/100 avg, {final_word_count} body words'})}\n\n"
    
    # Save
    try:
        essay = Essay(
            user_id=current_user.id,
            title=instructions[:100],
            instructions=instructions,
            final_content=best_draft,
            final_score=final_avg,
            rounds_used=len(all_scores['quill']),
            word_count_limit=word_count
        )
        db.session.add(essay)
        db.session.commit()
        
        yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': final_avg})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'DB error: {str(e)}'})}\n\n"

# TOOL C: The Oracle (AI Detection)
def check_ai_content(text):
    """Use GPT-4 to detect AI-generated content"""
    prompt = f"""You are an AI content detection expert. Analyze the following text to determine if it was written by AI or a human.

TEXT TO ANALYZE:
{text}

Look for AI indicators:
- Repetitive phrasing patterns
- Overly formal or perfect grammar
- Lack of personal voice
- Generic transitions
- AI buzzwords (delve, tapestry, multifaceted, landscape, etc.)

Return JSON format:
{{
    "ai_probability": 0-100,
    "indicators": ["list of specific AI indicators found"],
    "verdict": "HUMAN/LIKELY_HUMAN/UNCERTAIN/LIKELY_AI/AI",
    "explanation": "brief explanation"
}}"""

    try:
        response = call_gpt4(prompt)
        return response
    except Exception as e:
        return json.dumps({"error": str(e), "ai_probability": 0, "verdict": "ERROR"})

# TOOL D: The Grader - 3-Round Consensus Debate
def grade_assignment(brief_text, essay_text):
    """Grade assignment using 3-round consensus debate among 3 professors"""
    
    # ROUND 1: Independent Initial Grading
    print("🎓 ROUND 1: Independent Grading...")
    
    # Prof. Quill - Content Analysis (Claude)
    quill_prompt_r1 = f"""You are Prof. Quill, an expert academic evaluator. Analyze this student essay against the assignment brief.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate independently:
1. How well does it address the brief?
2. Content quality and depth
3. Structure and organization

Provide:
- A score (0-100)
- Detailed feedback (2-3 paragraphs)
- Key strengths and weaknesses"""

    # Dr. Strict - Technical Grading (GPT-4o)
    strict_prompt_r1 = f"""You are Dr. Strict, a harsh academic grader. Grade this essay strictly and independently.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate independently:
1. Grammar and writing quality
2. Citation and referencing
3. Academic rigor

Provide:
- A score (0-100)
- Detailed feedback identifying specific weaknesses
- Technical issues found"""

    # Dean Logic - Overall Assessment (Gemini)
    logic_prompt_r1 = f"""You are Dean Logic, the final academic authority. Provide an independent overall assessment.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Provide independently:
- Overall score (0-100)
- Key strengths
- Critical improvements needed
- Holistic evaluation"""

    try:
        # Round 1: Get initial independent grades
        quill_response_r1 = call_claude(quill_prompt_r1)
        quill_score_r1 = extract_score(quill_response_r1)
        
        strict_response_r1 = call_gpt4(strict_prompt_r1)
        strict_score_r1 = extract_score(strict_response_r1)
        
        logic_response_r1 = call_gemini(logic_prompt_r1)
        logic_score_r1 = extract_score(logic_response_r1)
        
        print(f"Round 1 Scores - Quill: {quill_score_r1}, Strict: {strict_score_r1}, Logic: {logic_score_r1}")
        
        # ROUND 2: Consensus Phase - Professors read each other's feedback
        print("🎓 ROUND 2: Consensus Phase...")
        
        quill_prompt_r2 = f"""You are Prof. Quill. You've read the feedback from Dr. Strict and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        strict_prompt_r2 = f"""You are Dr. Strict. You've read the feedback from Prof. Quill and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        logic_prompt_r2 = f"""You are Dean Logic. You've read the feedback from Prof. Quill and Dr. Strict.

YOUR INITIAL ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        # Round 2: Get revised consensus grades
        quill_response_r2 = call_claude(quill_prompt_r2)
        quill_score_r2 = extract_score(quill_response_r2)
        
        strict_response_r2 = call_gpt4(strict_prompt_r2)
        strict_score_r2 = extract_score(strict_response_r2)
        
        logic_response_r2 = call_gemini(logic_prompt_r2)
        logic_score_r2 = extract_score(logic_response_r2)
        
        print(f"Round 2 Scores - Quill: {quill_score_r2}, Strict: {strict_score_r2}, Logic: {logic_score_r2}")
        
        # ROUND 3: Final Calculation
        print("🎓 ROUND 3: Final Calculation...")
        
        # Use Round 2 scores (post-consensus) for final average
        average_score = round((quill_score_r2 + strict_score_r2 + logic_score_r2) / 3)
        
        result = {
            "average_score": average_score,
            "quill_score": quill_score_r2,
            "quill_feedback": quill_response_r2,
            "quill_initial_score": quill_score_r1,
            "strict_score": strict_score_r2,
            "strict_feedback": strict_response_r2,
            "strict_initial_score": strict_score_r1,
            "logic_score": logic_score_r2,
            "logic_feedback": logic_response_r2,
            "logic_initial_score": logic_score_r1,
            "debate_summary": f"Round 1: Quill {quill_score_r1}, Strict {strict_score_r1}, Logic {logic_score_r1}. Round 2 (Consensus): Quill {quill_score_r2}, Strict {strict_score_r2}, Logic {logic_score_r2}. Final Average: {average_score}/100"
        }
        
        print(f"✅ Final Average Score: {average_score}/100")
        
        return json.dumps(result)
    except Exception as e:
        print(f"❌ Error in grading: {str(e)}")
        return json.dumps({"error": str(e), "average_score": 0})

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
            return redirect(url_for('dashboard'))
        flash('Invalid credentials', 'error')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        
        # Check if passwords match
        if password != confirm_password:
            flash('Passwords do not match!', 'error')
            return redirect(url_for('register'))
        
        # Check for existing username
        if User.query.filter_by(username=username).first():
            flash('Username already exists. Please choose a different username.', 'error')
            return redirect(url_for('register'))
        
        # Check for existing email
        if User.query.filter_by(email=email).first():
            flash('Email already registered. Please use a different email or login.', 'error')
            return redirect(url_for('register'))
        
        try:
            user = User(username=username, email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            
            # Send mock welcome email
            email_body = f"""
Welcome to The Academic Board, {username}!

Your account has been successfully created.

Get started with our premium AI-powered academic tools:
- The Architect: AI Essay Writer (Now with 4-Agent Consensus!)
- The Detective: Plagiarism Checker
- The Oracle: AI Content Detector
- The Grader: Assignment Marking

Login now at: {DOMAIN}

Best regards,
The Academic Board Team
            """
            send_email_mock(email, "Welcome to The Academic Board!", email_body)
            
            flash('Welcome aboard! (Check console for mock email)', 'success')
            return redirect(url_for('login'))
        except Exception as e:
            db.session.rollback()
            flash('Registration failed. Please try again.', 'error')
            print(f"Registration error: {str(e)}")
            return redirect(url_for('register'))
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', user=current_user)

@app.route('/tools/writer')
@login_required
def tool_writer():
    return render_template('tool_writer.html', user=current_user)

@app.route('/tools/plagiarism')
@login_required
def tool_plagiarism():
    return render_template('tool_plagiarism.html', user=current_user)

@app.route('/tools/detector')
@login_required
def tool_detector():
    return render_template('tool_detector.html', user=current_user)

@app.route('/tools/grader')
@login_required
def tool_grader():
    return render_template('tool_grader.html', user=current_user)

@app.route('/billing')
@login_required
def billing():
    payments = Payment.query.filter_by(user_id=current_user.id).order_by(Payment.created_at.desc()).all()
    return render_template('billing.html', payments=payments, user=current_user)

@app.route('/my-essays')
@login_required
def my_essays():
    essays = Essay.query.filter_by(user_id=current_user.id).order_by(Essay.created_at.desc()).all()
    return render_template('my_essays.html', essays=essays, user=current_user)

@app.route('/settings')
@login_required
def settings():
    return render_template('settings.html', user=current_user)

# API Routes
@app.route('/api/support', methods=['POST'])
@login_required
def support_chat():
    """The Concierge - AI Support Bot with Error Handling"""
    try:
        data = request.get_json()
        user_message = data.get('message', '')
        
        if not user_message:
            return jsonify({'reply': 'Please provide a message.'}), 400
        
        # Check if OpenAI API key exists
        if not os.getenv('OPENAI_API_KEY'):
            return jsonify({'reply': 'System Error: Contact Admin - OpenAI API key not configured.'}), 500
        
        system_prompt = """You are The Concierge, the helpful support agent for The Academic Board.

PRICING:
- Essay Writer: £10 per 1,000 words (1,000-12,000 words range)
- Plagiarism Check: £10
- AI Detector: £10
- Assignment Grader: £10

TOOLS:
- The Architect: AI essay writer with 4-agent consensus system (Prof. Quill, Dr. Strict, Dean Logic, Chancellor GPT)
- The Detective: Plagiarism checker with PDF reports
- The Oracle: AI content detector
- The Grader: Strict assignment marking (now £10)

Be polite, concise, and helpful. Troubleshoot errors and explain features."""
        
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ],
            temperature=0.7,
            max_tokens=300
        )
        
        bot_reply = response.choices[0].message.content
        return jsonify({'reply': bot_reply})
        
    except openai.AuthenticationError:
        return jsonify({'reply': 'System Error: Contact Admin - Invalid API credentials.'}), 500
    except openai.RateLimitError:
        return jsonify({'reply': 'System Error: Contact Admin - Rate limit exceeded.'}), 500
    except openai.APIError as e:
        return jsonify({'reply': f'System Error: Contact Admin - API error occurred.'}), 500
    except Exception as e:
        return jsonify({'reply': 'System Error: Contact Admin - An unexpected error occurred.'}), 500

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    """
    MEMORY-OPTIMIZED file upload handler for 25MB files.
    Processes files efficiently to avoid RAM crashes on limited server.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Supported: PDF, DOCX, TXT (max 25MB each)'}), 400
    
    filepath = None
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save file
        file.save(filepath)
        
        # MEMORY-EFFICIENT: Extract text with limits (100 pages PDF, 300 paragraphs DOCX)
        text = extract_text_from_file(filepath, filename)
        
        # Clean up immediately to free RAM
        os.remove(filepath)
        filepath = None
        
        # Limit extracted text for frontend display (max 100KB)
        # Full text is already processed, this is just for preview
        if len(text) > 100000:
            text = text[:100000] + "\n\n[Preview truncated. Full content will be used in essay generation.]"
        
        print(f"✅ File processed successfully: {filename}")
        return jsonify({'success': True, 'text': text})
        
    except Exception as e:
        # Clean up on error
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        print(f"❌ Upload error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/create-checkout-session', methods=['POST'])
@login_required
def create_checkout_session():
    try:
        data = request.get_json()
        tool_type = data.get('tool_type')
        
        print(f"🔔 Checkout session requested for tool: {tool_type}")
        
        if tool_type not in PRICING:
            return jsonify({'error': 'Invalid tool'}), 400
        
        # Calculate amount based on tool type
        if tool_type == 'writer':
            # Dynamic pricing for writer
            word_count = data.get('word_count', 2000)
            amount = int((word_count / 1000) * PRICING['writer']['price_per_1k'])
            product_name = f'Essay Writing ({word_count:,} words)'
            print(f"💰 Writer pricing: {word_count} words = £{amount/100}")
        else:
            # Fixed pricing for other tools
            amount = PRICING[tool_type]['price']
            product_names = {
                'plagiarism': 'Plagiarism Check',
                'ai_check': 'AI Detection',
                'grader': 'Assignment Grading'
            }
            product_name = product_names.get(tool_type, 'Academic Tool')
        
        session['pending_task'] = {'tool_type': tool_type, 'data': data}
        
        # Hardcode Render domain for success_url
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price_data': {
                    'currency': 'gbp',
                    'unit_amount': amount,
                    'product_data': {'name': product_name},
                },
                'quantity': 1,
            }],
            mode='payment',
            allow_promotion_codes=True,
            success_url='https://aiassignment-x631.onrender.com/payment-success?session_id={CHECKOUT_SESSION_ID}&tool=' + tool_type,
            cancel_url='https://aiassignment-x631.onrender.com/dashboard',
            client_reference_id=str(current_user.id),
            customer_email=current_user.email,
        )
        
        print(f"✅ Stripe session created: {checkout_session.id}")
        return jsonify({'id': checkout_session.id})
    except Exception as e:
        print(f"❌ Checkout error: {str(e)}")
        return jsonify({'error': str(e)}), 403

@app.route('/payment-success')
@login_required
def payment_success():
    session_id = request.args.get('session_id')
    tool_type = request.args.get('tool')
    
    if session_id:
        try:
            checkout_session = stripe.checkout.Session.retrieve(session_id)
            
            if checkout_session.payment_status == 'paid':
                # Record payment
                payment = Payment(
                    user_id=current_user.id,
                    amount=checkout_session.amount_total,
                    tool_type=tool_type,
                    stripe_session_id=session_id
                )
                db.session.add(payment)
                db.session.commit()
                
                flash('Payment successful!', 'success')
                
                # Redirect to appropriate tool
                tool_routes = {
                    'writer': 'tool_writer',
                    'plagiarism': 'tool_plagiarism',
                    'ai_check': 'tool_detector',
                    'grader': 'tool_grader'
                }
                return redirect(url_for(tool_routes.get(tool_type, 'dashboard')) + '?paid=true')
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('dashboard'))

@app.route('/generate-essay', methods=['POST'])
@login_required
def generate_essay():
    instructions = request.form.get('instructions')
    word_count = int(request.form.get('word_count', 2000))
    
    if not instructions:
        return jsonify({'error': 'Instructions required'}), 400
    
    return Response(
        stream_with_context(generate_essay_stream(instructions, word_count)),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'X-Accel-Buffering': 'no'
        }
    )

@app.route('/check-plagiarism', methods=['POST'])
@login_required
def check_plagiarism():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        # Step 1: Hunter (Gemini)
        hunter_findings = plagiarism_hunter_gemini(text)
        
        # Step 2: Analyst (GPT-4o)
        analyst_findings = plagiarism_analyst_gpt(text, hunter_findings)
        
        # Step 3: Reporter (Claude) - Generate PDF
        pdf_filename = plagiarism_reporter_claude(text, hunter_findings, analyst_findings, current_user.id)
        
        # Save report
        report = Report(
            user_id=current_user.id,
            tool_type='plagiarism',
            title='Plagiarism Check',
            result_data=analyst_findings,
            file_path=pdf_filename
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'pdf_url': f'/download-report/{report.id}'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/check-ai', methods=['POST'])
@login_required
def check_ai():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        analysis = check_ai_content(text)
        
        report = Report(
            user_id=current_user.id,
            tool_type='ai_check',
            title='AI Detection',
            result_data=analysis
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'analysis': analysis
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/grade-assignment', methods=['POST'])
@login_required
def grade_assignment_route():
    brief = request.form.get('brief')
    essay = request.form.get('essay')
    
    if not brief or not essay:
        return jsonify({'error': 'Both brief and essay required'}), 400
    
    try:
        grading = grade_assignment(brief, essay)
        
        report = Report(
            user_id=current_user.id,
            tool_type='grader',
            title='Assignment Grading',
            result_data=grading
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'grading': grading
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download-report/<int:report_id>')
@login_required
def download_report(report_id):
    report = Report.query.get_or_404(report_id)
    
    if report.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))
    
    if report.file_path:
        filepath = os.path.join(app.config['REPORTS_FOLDER'], report.file_path)
        return send_file(filepath, as_attachment=True)
    
    return jsonify({'error': 'No file'}), 404

@app.route('/download-essay/<int:essay_id>')
@login_required
def download_essay(essay_id):
    essay = Essay.query.get_or_404(essay_id)
    
    if essay.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))
    
    doc = Document()
    title = doc.add_heading(essay.title, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Generated: {essay.created_at.strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Score: {essay.final_score}/100")
    doc.add_paragraph("")
    doc.add_heading('Essay', 1)
    doc.add_paragraph(essay.final_content)
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return Response(
        file_stream.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers={'Content-Disposition': f'attachment; filename=essay_{essay_id}.docx'}
    )

# Initialize database
with app.app_context():
    print("🔧 Creating database...")
    db.create_all()
    print("✅ Database ready!")
    
    if not User.query.filter_by(username='demo').first():
        demo = User(username='demo', email='demo@example.com')
        demo.set_password('demo123')
        db.session.add(demo)
        db.session.commit()
        print("✅ Demo user created")

if __name__ == '__main__':
    # Run with increased timeout for long-running AI operations
    # CRITICAL: Gunicorn timeout MUST be set to 300s in Render Start Command: gunicorn --timeout 300 app:app
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
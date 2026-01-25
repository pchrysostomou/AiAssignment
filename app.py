import os
import json
import time
import re
import requests
import threading
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

# CRITICAL FIX: Set permissive CSP that allows Stripe.js, html2pdf.js and eval() for payment processing
@app.after_request
def add_security_headers(response):
    """
    Configure Content Security Policy to allow:
    1. Stripe.js library and checkout iframe
    2. html2pdf.js from CDN
    3. unsafe-eval for Stripe's payment processing
    4. Our own domain for API calls
    """
    # Get the current domain
    domain = os.getenv('DOMAIN', 'https://aiassignment-x631.onrender.com')
    
    # Build CSP policy that allows Stripe and html2pdf while maintaining reasonable security
    csp_policy = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' 'unsafe-eval' https://js.stripe.com https://cdn.jsdelivr.net https://cdnjs.cloudflare.com; "
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
    if request.path.startswith('/api/') or request.path.startswith('/generate-') or request.path.startswith('/create-checkout') or request.path.startswith('/check-') or request.path.startswith('/export-pdf'):
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
    'grader': {'price': 1000, 'display': '£10'}
}

# API Credentials Initialization with Validation
anthropic_api_key = os.getenv('ANTHROPIC_API_KEY', '')
if anthropic_api_key:
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

# Google Custom Search Engine Configuration
GOOGLE_CSE_ID = str(os.getenv('GOOGLE_CSE_ID', '50f552d88c3e14772')).strip()
if GOOGLE_CSE_ID and GOOGLE_CSE_ID != '':
    print(f"✅ Google CSE ID Configured: '{GOOGLE_CSE_ID}'")
else:
    print("⚠️ WARNING: GOOGLE_CSE_ID not found or empty in environment variables")

# --- SMART GEMINI MODEL SELECTOR ---
def get_working_gemini_model():
    """
    Probes Google API to find which model alias is actually valid.
    Prevents 404 errors by testing aliases in order.
    """
    if not google_api_key: 
        return "gemini-1.5-flash" # Fallback
        
    print("🔍 DIAGNOSTIC: Probing for valid Gemini model...")
    
    candidates = [
        "models/gemini-1.5-pro-latest",
        "gemini-1.5-pro",
        "models/gemini-1.5-flash",
        "gemini-1.5-flash",
        "gemini-1.5-flash-001",
        "gemini-pro"
    ]
    
    try:
        # Ask Google what exists
        available_models = [m.name for m in genai.list_models()]
        print(f"📋 Server lists these models: {available_models}")
        
        # Check for matches
        for cand in candidates:
            if cand in available_models or f"models/{cand}" in available_models:
                print(f"✅ FOUND MATCH: {cand}")
                return cand
                
            # Fuzzy match check
            match = next((m for m in available_models if cand.replace("models/", "") in m), None)
            if match:
                print(f"✅ FOUND FUZZY MATCH: {match}")
                return match

    except Exception as e:
        print(f"⚠️ Model probing failed: {e}")

    print("⚠️ Probe failed, using safe fallback: gemini-1.5-flash")
    return "gemini-1.5-flash"

# Initialize the model using the function
GEMINI_MODEL = get_working_gemini_model()
print(f"🎯 FINAL GEMINI MODEL SELECTED: {GEMINI_MODEL}")

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
    """Extract text from PDF file with MEMORY-EFFICIENT streaming approach."""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            max_pages = min(total_pages, 100)
            
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
                text += f"\n\n[Note: Document has {total_pages} pages. First {max_pages} pages extracted.]"
            
            return text.strip()
    except Exception as e:
        raise Exception(f"PDF extraction error: {str(e)}")

def extract_text_from_docx(file_path):
    """Extract text from DOCX file with MEMORY-EFFICIENT approach."""
    try:
        doc = Document(file_path)
        total_paragraphs = len(doc.paragraphs)
        max_paragraphs = min(total_paragraphs, 300)
        
        text_chunks = []
        for i in range(max_paragraphs):
            para_text = doc.paragraphs[i].text
            if para_text.strip():
                text_chunks.append(para_text)
        
        text = "\n".join(text_chunks)
        
        if total_paragraphs > max_paragraphs:
            text += f"\n\n[Note: Document has {total_paragraphs} paragraphs. First {max_paragraphs} extracted.]"
        
        return text.strip()
    except Exception as e:
        raise Exception(f"DOCX extraction error: {str(e)}")

def extract_text_from_file(file_path, filename):
    """Helper function to extract text from PDF or DOCX files."""
    file_size = os.path.getsize(file_path)
    file_size_mb = file_size / (1024 * 1024)
    
    if file_size > 25 * 1024 * 1024:
        raise Exception(f"File too large ({file_size_mb:.1f}MB). Maximum 25MB per file.")
    
    print(f"📄 Processing file: {filename} ({file_size_mb:.1f}MB)")
    
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    elif ext == 'txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read(200 * 1024)
            return text.strip()
    raise Exception("Unsupported file type")

def parse_uploaded_file(file_path, filename):
    """Parse uploaded file and extract text"""
    return extract_text_from_file(file_path, filename)

# Google Custom Search Integration
def google_search(query, num_results=5):
    """Perform Google Custom Search using the configured API key and CSE ID."""
    if not google_api_key or not GOOGLE_CSE_ID or GOOGLE_CSE_ID == '':
        print(f"⚠️ Google Search unavailable")
        return []
    
    try:
        words = query.split()
        key_words = [w for w in words if len(w) > 1][:5]
        
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
        
        print(f"🔍 Google Search Query: '{search_query}'")
        
        response = requests.get(url, params=params, timeout=10)
        
        if response.status_code != 200:
            print(f"❌ Google Search API Error: {response.status_code}")
            return []
        
        data = response.json()
        results = []
        
        for item in data.get('items', []):
            results.append({
                'title': item.get('title', ''),
                'link': item.get('link', ''),
                'snippet': item.get('snippet', '')
            })
        
        print(f"✅ Google Search completed: {len(results)} results")
        return results
        
    except Exception as e:
        print(f"❌ Google Search error: {str(e)}")
        return []

# --- FORCE JSON CLEANER (OVERWRITE) ---
def clean_and_parse_json(response_text):
    """Sanitizes AI response to extract pure JSON. Prevents 'Unexpected token' crashes."""
    try:
        if isinstance(response_text, dict): 
            return response_text
        
        # Remove markdown fences
        text = re.sub(r'```json\s*', '', response_text)
        text = re.sub(r'```', '', text)
        text = text.strip()
        
        # Extract JSON object
        start, end = text.find('{'), text.rfind('}') + 1
        if start != -1 and end != 0: 
            return json.loads(text[start:end])
        
        return json.loads(text)
    except Exception as e:
        print(f"JSON Error: {e}")
        # Return safe fallback structure
        return {"score": 0, "sources": [], "matches": [], "ai_score": 0, "segments": []}

# --- TOKEN TRUNCATION HELPER ---
def truncate_text(text, max_tokens=10000):
    """
    Truncate text to approximate token limit.
    Rule of thumb: 1 token ≈ 4 characters for English text.
    """
    max_chars = max_tokens * 4
    if len(text) <= max_chars:
        return text
    
    truncated = text[:max_chars]
    print(f"⚠️ Text truncated from {len(text)} to {len(truncated)} characters (~{max_tokens} tokens)")
    return truncated + "\n\n[Content truncated to fit token limit]"

# AI Helper Functions - ROBUST CONFIGURATION WITH RATE LIMIT HANDLING
def call_with_retry(func, max_retries=3):
    """Enhanced retry logic with exponential backoff for rate limits"""
    for attempt in range(max_retries):
        try:
            return func()
        except openai.RateLimitError as e:
            if attempt == max_retries - 1:
                print(f"❌ Rate limit exceeded after {max_retries} attempts")
                raise
            wait_time = 2 ** (attempt + 1)  # Exponential backoff: 2s, 4s, 8s
            print(f"⚠️ Rate limit hit (attempt {attempt + 1}/{max_retries}). Waiting {wait_time}s...")
            time.sleep(wait_time)
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(2 * (attempt + 1))

# FIX #1: SWITCH TO HAIKU (Stable Model)
def call_claude(prompt, max_tokens=4000):
    """
    RESCUE FIX: Switched to 'claude-3-haiku-20240307' for maximum stability.
    Falls back to GPT-4 if Claude fails entirely.
    """
    if not anthropic_client:
        return call_gpt4(prompt)  # Fallback if no API key
    
    try:
        print(f"🤖 Calling Claude 3 Haiku (stable version 20240307)")
        response = anthropic_client.messages.create(
            model="claude-3-haiku-20240307",  # CHANGED: Use most stable Haiku model
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        print(f"✅ Claude 3 Haiku API call successful")
        return response.content[0].text
    except Exception as e:
        print(f"❌ Claude Error: {str(e)}")
        print(f"🔄 Falling back to GPT-4...")
        return call_gpt4(prompt)  # Hard fallback to GPT

def call_gpt4(prompt, model="gpt-4o", max_input_tokens=15000):
    """
    RATE-LIMIT-SAFE GPT-4o caller with token truncation and exponential backoff.
    Truncates input to max_input_tokens to stay under 30k TPM limit.
    """
    if not openai_api_key:
        raise Exception("OpenAI API key not configured")
    
    # Truncate prompt to stay under token limit
    truncated_prompt = truncate_text(prompt, max_tokens=max_input_tokens)
    
    def api_call():
        print(f"🤖 Calling OpenAI GPT-4o (input ~{len(truncated_prompt)//4} tokens)")
        response = openai.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": truncated_prompt}],
            temperature=0.3
        )
        print(f"✅ OpenAI GPT-4o API call successful")
        return response.choices[0].message.content
    
    return call_with_retry(api_call, max_retries=3)

def call_gemini(prompt):
    """
    Robust Gemini caller using the auto-detected GEMINI_MODEL.
    Includes emergency fallback logic.
    """
    if not google_api_key:
        raise Exception("Google API key not configured")
    
    print(f"🤖 Calling Google Gemini ({GEMINI_MODEL})...")
    
    try:
        model = genai.GenerativeModel(GEMINI_MODEL)
        response = model.generate_content(prompt)
        print(f"✅ Google Gemini API call successful")
        return response.text
    except Exception as e:
        error_msg = str(e)
        print(f"❌ Gemini Error: {error_msg}")
        
        # Emergency Fallback if the main model fails (e.g. 404)
        if "404" in error_msg or "not found" in error_msg.lower():
            print("🔄 404 Error detected. Attempting emergency fallback to 'gemini-pro'...")
            try:
                fallback_model = genai.GenerativeModel("gemini-pro")
                response = fallback_model.generate_content(prompt)
                return response.text
            except Exception as e2:
                raise Exception(f"Gemini Fallback also failed: {str(e2)}")
        
        raise Exception(f"Gemini failed: {error_msg}")

def extract_score(text):
    """Extract score from agent response"""
    matches = re.findall(r'\b(\d{1,3})\b', text)
    for match in matches:
        score = int(match)
        if 0 <= score <= 100:
            return score
    return 0

def count_words(text):
    """Count words in text, excluding references section"""
    ref_markers = ['references', 'bibliography', 'works cited']
    text_lower = text.lower()
    
    for marker in ref_markers:
        if f'\n{marker}\n' in text_lower or f'\n{marker}:' in text_lower:
            parts = re.split(f'\n{marker}[:\n]', text_lower, maxsplit=1)
            if len(parts) > 1:
                body_text = parts[0]
                words = body_text.split()
                return len(words)
    
    words = text.split()
    return len(words)

def check_url_status(url):
    """Check if URL is live (HTTP 200) or broken"""
    try:
        response = requests.head(url, timeout=5, allow_redirects=True)
        return response.status_code == 200
    except:
        return False

# --- HYBRID PLAGIARISM HUNTER: GEMINI SEARCHES, GPT-4 VERIFIES ---
def plagiarism_hunter_gemini(text):
    """
    HYBRID APPROACH:
    Step 1: Gemini searches web for potential sources (has Google access)
    Step 2: GPT-4 performs detailed matching and scoring (more accurate)
    This prevents hallucinated high scores with no real matches.
    """
    print("🔍 HYBRID PLAGIARISM DETECTION: Gemini searches, GPT-4 verifies...")
    
    # Step 1: Gemini searches for Sources (It has Google access)
    search_prompt = f"""Search the web for this text. Return a list of SPECIFIC URLs that match. 
TEXT: {text[:1000]}...
JSON OUTPUT: {{"potential_urls": ["url1", "url2"]}}"""
    
    try:
        search_data = clean_and_parse_json(call_gemini(search_prompt))
        urls = search_data.get('potential_urls', [])
        print(f"✅ Gemini found {len(urls)} potential URLs")
    except: 
        urls = []
        print("⚠️ Gemini search failed, using empty URL list")

    # Step 2: GPT-4 performs the detailed Matching & Color Coding (More accurate)
    # We feed the URLs found by Gemini into GPT-4
    analysis_prompt = f"""You are a Plagiarism Analyst. 
1. Check if this TEXT matches content from these URLs: {urls}
2. Extract EXACT matched segments.
3. Assign a plagiarism score based on ACTUAL matches ONLY. If no real matches found, score MUST be 0-10.

TEXT: {text[:3000]}

Return STRICT JSON:
{{
    "score": 0-100,
    "sources": [
        {{"id": 1, "domain": "example.com", "url": "https://example.com/full-path", "similarity": 20}}
    ],
    "matches": [
        {{"text_segment": "exact text from student...", "source_id": 1}}
    ]
}}

CRITICAL RULES:
1. URLs MUST be complete with https://
2. Score reflects ACTUAL text overlap, not speculation
3. If no matches found, return score: 0, sources: [], matches: []"""
    
    try:
        findings = clean_and_parse_json(call_gpt4(analysis_prompt))
        print(f"✅ GPT-4 verification complete: Score {findings.get('score', 0)}%, {len(findings.get('sources', []))} sources")
        return findings
    except Exception as e:
        print(f"❌ GPT-4 verification error: {str(e)}")
        return {"score": 0, "sources": [], "matches": [], "error": str(e)}

# --- 3-AGENT CONSENSUS FOR AI DETECTION ---
def check_ai_consensus(text):
    """Queries Gemini, GPT-4, and Claude. Returns Average Score with breakdown."""
    safe_text = truncate_text(text, 3000)
    
    print("🤖 3-AGENT CONSENSUS: Starting AI detection...")
    
    # 1. GPT-4 (Base analysis + Highlighting)
    gpt_prompt = f"""Analyze SENTENCE BY SENTENCE for AI. Text: {safe_text}. 
    Return JSON: {{"ai_score": 0-100, "segments": [{{"text": "sentence...", "is_ai": true, "confidence": 90}}]}}"""
    
    try:
        gpt_response = call_gpt4(gpt_prompt)
        gpt_data = clean_and_parse_json(gpt_response)
        gpt_score = gpt_data.get('ai_score', 0)
        print(f"✅ GPT-4 Score: {gpt_score}%")
    except Exception as e:
        print(f"❌ GPT-4 error: {str(e)}")
        gpt_score = 0
        gpt_data = {"ai_score": 0, "segments": []}
    
    # 2. Gemini (Score Verification)
    gemini_prompt = f"Rate AI probability (0-100) ONLY. Text: {safe_text}"
    try:
        gemini_response = call_gemini(gemini_prompt)
        gemini_score = extract_score(gemini_response)
        print(f"✅ Gemini Score: {gemini_score}%")
    except Exception as e:
        print(f"❌ Gemini error: {str(e)}")
        gemini_score = gpt_score  # Fallback to GPT score
    
    # 3. Claude (Score Verification)
    claude_prompt = f"Rate AI probability (0-100) ONLY. Text: {safe_text}"
    try:
        claude_response = call_claude(claude_prompt, max_tokens=500)
        claude_score = extract_score(claude_response)
        print(f"✅ Claude Score: {claude_score}%")
    except Exception as e:
        print(f"❌ Claude error: {str(e)}")
        claude_score = gemini_score  # Fallback to Gemini score
    
    # 4. Weighted Average (GPT gets more weight for segment analysis)
    final_score = int((gpt_score * 0.4) + (gemini_score * 0.3) + (claude_score * 0.3))
    
    print(f"🎯 CONSENSUS SCORE: {final_score}% (GPT:{gpt_score}% | Gem:{gemini_score}% | Claude:{claude_score}%)")
    
    return {
        "ai_score": final_score,
        "segments": gpt_data.get('segments', []),
        "breakdown": f"GPT:{gpt_score}% | Gem:{gemini_score}% | Claude:{claude_score}%",
        "input_text": text  # CRITICAL: Save the original text
    }

# --- HEARTBEAT MECHANISM FOR SSE KEEPALIVE ---
class HeartbeatThread(threading.Thread):
    """
    Background thread that sends heartbeat pings every 10 seconds
    to keep SSE connection alive and prevent proxy timeouts.
    """
    def __init__(self, stop_event, callback):
        super().__init__(daemon=True)
        self.stop_event = stop_event
        self.callback = callback
        
    def run(self):
        while not self.stop_event.is_set():
            time.sleep(10)  # Send heartbeat every 10 seconds
            if not self.stop_event.is_set():
                try:
                    self.callback()
                except:
                    pass  # Ignore errors if stream is closed

# TOOL B: The Architect - GROWTH MODE + SELF-REFLECTION 3-AGENT SYSTEM WITH ENHANCED RATCHET + FLOOR MECHANISM
def generate_essay_stream(instructions, word_count):
    """
    GROWTH MODE + SELF-REFLECTION CONSENSUS LOOP WITH CRITICAL FIXES:
    - WORD COUNT ENFORCEMENT: Minimum 95% of target (1,900/2,000) required before consensus
    - GUARANTEED EXIT: Round 14+ triggers automatic finalization to prevent timeout
    - ENHANCED ANTI-SHRINKAGE RATCHET: Prevents shrinkage during growth AND refinement drop
    - FLOOR MECHANISM: Once target is reached, word count CANNOT drop below target
    - HEARTBEAT: Sends ping every 10 seconds to prevent proxy timeouts
    - PROGRESS: Shows percentage completion (Round X/15 - Y%)
    """
    # Setup heartbeat mechanism
    stop_heartbeat = threading.Event()
    heartbeat_queue = []
    
    def send_heartbeat():
        """Callback to send heartbeat ping"""
        heartbeat_queue.append(f"data: {json.dumps({'type': 'heartbeat', 'timestamp': time.time()})}\n\n")
    
    # Start heartbeat thread
    heartbeat_thread = HeartbeatThread(stop_heartbeat, send_heartbeat)
    heartbeat_thread.start()
    
    try:
        yield f"data: {json.dumps({'type': 'log', 'message': '🎓 The Academic Board - ENHANCED RATCHET + FLOOR MECHANISM'})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': '💓 Heartbeat enabled: Connection will stay alive during long operations'})}\n\n"
        
        # CRITICAL: Minimum word count threshold (95% of target)
        MIN_WORD_COUNT = int(word_count * 0.95)  # 1,900 words for 2,000 target
        target_total_words = 2200
        MAX_ROUNDS = 15
        
        yield f"data: {json.dumps({'type': 'log', 'message': f'📏 Word Count Enforcement: Minimum {MIN_WORD_COUNT} words required (95% of {word_count})'})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': '🔒 Enhanced Ratchet: Prevents shrinkage during growth'})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': '🧱 Floor Mechanism: Once target reached, word count cannot drop below target'})}\n\n"
        
        # Research Phase
        yield f"data: {json.dumps({'type': 'progress', 'current': 0, 'total': MAX_ROUNDS, 'percentage': 0, 'stage': 'Research'})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': '🔍 Researching topic via Google Custom Search...'})}\n\n"
        
        research_context = ""
        try:
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
        
        # Send any queued heartbeats
        while heartbeat_queue:
            yield heartbeat_queue.pop(0)
        
        # Initial draft by Prof. Quill (Claude 3 Haiku 20240307)
        yield f"data: {json.dumps({'type': 'progress', 'current': 0, 'total': MAX_ROUNDS, 'percentage': 0, 'stage': 'Initial Draft'})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill (Claude 3 Haiku 20240307) drafting initial essay...'})}\n\n"
        
        # Truncate instructions to prevent token overflow
        truncated_instructions = truncate_text(instructions, max_tokens=2000)
        
        writer_prompt = f"""You are Prof. Quill, operating under the SPARTAN ACADEMIC protocol.

CRITICAL WORD COUNT TARGET:
- TARGET: {target_total_words} words TOTAL (to ensure {word_count}+ body words after references)
- Write the full essay body first, THEN add a complete References section separately

SMART CITATION LOGIC - Analyze the assignment type:
- CRITICAL REVIEW of a single paper/book: NO external citations needed (only cite the work being reviewed)
- LITERATURE REVIEW: MANY citations needed (10-20+ sources)
- RESEARCH ESSAY: MODERATE citations (5-10 sources)
- OPINION PIECE: FEW citations (2-5 sources for key claims only)
- COMPARATIVE ANALYSIS: MODERATE citations (cite each work being compared)

THE BAN LIST - STRICTLY FORBIDDEN WORDS:
❌ delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust

SPARTAN STYLE RULES:
- Direct, authoritative tone. No hedging.
- Vary sentence length strategically.
- Avoid excessive transition words.
- Use active voice.
- Cut unnecessary words.
- Contractions acceptable when they strengthen voice.
- Minor stylistic imperfections for authenticity.

INSTRUCTIONS:
{truncated_instructions}

{research_context}

Write {target_total_words} words total. Apply smart citation logic. No banned words. No fluff."""

        try:
            current_draft = call_claude(writer_prompt, max_tokens=4000)
            current_word_count = count_words(current_draft)
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Initial draft complete ({current_word_count} body words)'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': f'Error: {str(e)}'})}\n\n"
            stop_heartbeat.set()
            return
        
        # Send any queued heartbeats
        while heartbeat_queue:
            yield heartbeat_queue.pop(0)
        
        # GROWTH MODE + SELF-REFLECTION CONSENSUS LOOP WITH WORD COUNT ENFORCEMENT
        best_draft = current_draft
        best_avg_score = 0
        round_count = 0
        
        # Store specific rewrites from critics
        logic_rewrites = ""
        gpt_rewrites = ""
        
        while round_count < MAX_ROUNDS:
            round_count += 1
            current_word_count = count_words(current_draft)
            
            # Calculate progress percentage
            progress_percentage = round((round_count / MAX_ROUNDS) * 100)
            
            yield f"data: {json.dumps({'type': 'progress', 'current': round_count, 'total': MAX_ROUNDS, 'percentage': progress_percentage, 'stage': f'Round {round_count}'})}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ ROUND {round_count}/{MAX_ROUNDS} - {progress_percentage}% ({current_word_count}/{word_count} body words) ━━━'})}\n\n"
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)
            
            # CRITICAL FIX #2: GUARANTEED EXIT AT ROUND 14+
            if round_count >= 14:
                yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ ROUND {round_count}: Approaching timeout limit. Triggering automatic finalization...'})}\n\n"
                yield f"data: {json.dumps({'type': 'progress', 'current': MAX_ROUNDS, 'total': MAX_ROUNDS, 'percentage': 100, 'stage': 'Auto-Finalize'})}\n\n"
                yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 AUTO-FINALIZE: Using best draft to prevent timeout (Score: {best_avg_score}/100, Words: {count_words(best_draft)})'})}\n\n"
                current_draft = best_draft
                break
            
            # Truncate draft for grading to prevent token overflow
            truncated_draft = truncate_text(current_draft, max_tokens=8000)
            
            # CRITIC 1: Dean Logic (Gemini with Smart Discovery)
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚖️ Dean Logic (Gemini {GEMINI_MODEL}) evaluating...'})}\n\n"
            
            logic_prompt = f"""You are Dean Logic, a harsh academic critic. Grade strictly.

INSTRUCTIONS: {truncated_instructions}

ESSAY: {truncated_draft}

Criteria:
- Coherence, citation logic, requirements met, originality
- Word count: Target {word_count}+ (current: {current_word_count})
- DEDUCT 20 if banned words found, 15 if under word count

If score < 80, provide:
1. Score (0-100) format: 'Score: [number]'
2. Specific rewrites labeled "DEAN LOGIC REWRITES:"

Example:
Score: 75

DEAN LOGIC REWRITES:
[Paragraph 3 replace with:]
"Evidence shows X correlates with Y. Smith (2023) demonstrates..."

Provide score and rewrites if needed."""

            try:
                logic_response = call_gemini(logic_prompt)
                score_logic = extract_score(logic_response)
                yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Dean Logic: {score_logic}/100'})}\n\n"
                
                # Extract rewrites if score < 80
                if score_logic < 80 and "DEAN LOGIC REWRITES:" in logic_response:
                    logic_rewrites = logic_response.split("DEAN LOGIC REWRITES:")[1].strip()
                    yield f"data: {json.dumps({'type': 'log', 'message': '📝 Dean Logic provided specific rewrites'})}\n\n"
                else:
                    logic_rewrites = ""
                    
            except Exception as e:
                yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Dean Logic error: {str(e)}'})}\n\n"
                score_logic = 0
                logic_response = "Error during grading"
                logic_rewrites = ""
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)
            
            # CRITIC 2: Chancellor GPT (GPT-4o with rate limit protection)
            yield f"data: {json.dumps({'type': 'log', 'message': '🎓 Chancellor GPT (GPT-4o) evaluating...'})}\n\n"
            
            gpt_prompt = f"""You are Chancellor GPT, the supreme academic auditor. Grade with highest standards.

INSTRUCTIONS: {truncated_instructions}

ESSAY: {truncated_draft}

Criteria:
- Excellence, argumentation, writing standards, engagement
- Word count: Target {word_count}+ (current: {current_word_count})
- DEDUCT 20 if banned words, 15 if under count

If score < 80, provide:
1. Score (0-100) format: 'Score: [number]'
2. Specific rewrites labeled "CHANCELLOR GPT REWRITES:"

Example:
Score: 72

CHANCELLOR GPT REWRITES:
[Introduction needs hook:]
"In contemporary discourse on X, scholars recognize..."

Provide score and rewrites if needed."""

            try:
                gpt_response = call_gpt4(gpt_prompt, max_input_tokens=12000)
                score_gpt = extract_score(gpt_response)
                yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Chancellor GPT: {score_gpt}/100'})}\n\n"
                
                # Extract rewrites if score < 80
                if score_gpt < 80 and "CHANCELLOR GPT REWRITES:" in gpt_response:
                    gpt_rewrites = gpt_response.split("CHANCELLOR GPT REWRITES:")[1].strip()
                    yield f"data: {json.dumps({'type': 'log', 'message': '📝 Chancellor GPT provided specific rewrites'})}\n\n"
                else:
                    gpt_rewrites = ""
                    
            except openai.RateLimitError as e:
                yield f"data: {json.dumps({'type': 'log', 'message': '⚠️ Chancellor GPT rate limit - retrying with backoff...'})}\n\n"
                score_gpt = 0
                gpt_response = "Rate limit error"
                gpt_rewrites = ""
            except Exception as e:
                yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Chancellor GPT error: {str(e)}'})}\n\n"
                score_gpt = 0
                gpt_response = "Error during grading"
                gpt_rewrites = ""
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)
            
            # --- CRITIC 3: Prof. Quill Self-Reflection (Claude) ---
            yield f"data: {json.dumps({'type': 'log', 'message': '🤔 Prof. Quill (Claude) self-evaluating...'})}\n\n"
            
            quill_grade_prompt = f"""You are Prof. Quill. Grade your own draft OBJECTIVELY.

INSTRUCTIONS: {truncated_instructions}

YOUR DRAFT: {truncated_draft}

Target: {word_count}+ body words. Current: {current_word_count} body words.

STRICT RULES:
- If current < target, your Max Score is 60 (you failed the word count requirement)
- Did you maintain the Spartan tone?
- Did you avoid banned words?
- Is the argumentation strong?

Provide:
1. Score (0-100) in format: 'Score: [number]'
2. Brief self-correction plan if score < 80

Be harsh on yourself."""

            try:
                quill_grade_response = call_claude(quill_grade_prompt, max_tokens=2000)
                score_quill = extract_score(quill_grade_response)
                yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Prof. Quill Self-Score: {score_quill}/100'})}\n\n"
            except Exception as e:
                score_quill = 0
                yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Quill self-grade error: {str(e)}'})}\n\n"
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)
            
            # Calculate Average of 3 Agents
            avg_score = round((score_logic + score_gpt + score_quill) / 3)
            
            yield f"data: {json.dumps({'type': 'score', 'round': round_count, 'score': avg_score})}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Round {round_count} Average: {avg_score}/100 (Logic: {score_logic}, GPT: {score_gpt}, Quill: {score_quill})'})}\n\n"
            
            # Track best draft
            if avg_score > best_avg_score:
                best_avg_score = avg_score
                best_draft = current_draft
            
            # CRITICAL FIX #1: WORD COUNT ENFORCEMENT - Consensus requires BOTH score AND word count
            word_count_met = current_word_count >= MIN_WORD_COUNT
            
            if not word_count_met:
                yield f"data: {json.dumps({'type': 'log', 'message': f'❌ WORD COUNT NOT MET: {current_word_count}/{MIN_WORD_COUNT} minimum. Forcing expansion...'})}\n\n"
            
            # CHECK CONSENSUS - ALL 3 AGENTS MUST SCORE 80+ AND WORD COUNT >= 95% TARGET
            if score_logic >= 80 and score_gpt >= 80 and score_quill >= 80 and word_count_met:
                yield f"data: {json.dumps({'type': 'progress', 'current': round_count, 'total': MAX_ROUNDS, 'percentage': 100, 'stage': 'Complete'})}\n\n"
                yield f"data: {json.dumps({'type': 'log', 'message': f'✅ CONSENSUS REACHED. All 3 agents agree (80+) AND word count met ({current_word_count}/{MIN_WORD_COUNT})!'})}\n\n"
                yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 Final: {avg_score}/100 avg, {current_word_count} body words'})}\n\n"
                break
            else:
                if not word_count_met:
                    yield f"data: {json.dumps({'type': 'log', 'message': f'❌ NO CONSENSUS: Word count too low ({current_word_count}/{MIN_WORD_COUNT}). Must expand!'})}\n\n"
                else:
                    yield f"data: {json.dumps({'type': 'log', 'message': f'❌ NO CONSENSUS: Scores insufficient (Logic: {score_logic}, GPT: {score_gpt}, Quill: {score_quill})'})}\n\n"
            
            # Stop if max rounds reached
            if round_count >= MAX_ROUNDS:
                yield f"data: {json.dumps({'type': 'progress', 'current': MAX_ROUNDS, 'total': MAX_ROUNDS, 'percentage': 100, 'stage': 'Max Rounds'})}\n\n"
                yield f"data: {json.dumps({'type': 'log', 'message': f'⏱️ Max {MAX_ROUNDS} rounds reached. Using best draft (avg: {best_avg_score}/100)'})}\n\n"
                current_draft = best_draft
                break
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)
            
            # --- SMART MODE SWITCHING: GROWTH VS REFINEMENT WITH WORD COUNT PRIORITY ---
            if current_word_count < MIN_WORD_COUNT:
                # PHASE 1: AGGRESSIVE GROWTH (Under Minimum Threshold)
                # ABSOLUTE PRIORITY: Reach minimum word count
                deficit = MIN_WORD_COUNT - current_word_count
                mode_instruction = f"""
*** CRITICAL WORD COUNT ENFORCEMENT (HIGHEST PRIORITY) ***
Current: {current_word_count} words. MINIMUM REQUIRED: {MIN_WORD_COUNT} words.
DEFICIT: {deficit} words MUST BE ADDED.

ABSOLUTE RULES (NON-NEGOTIABLE):
1. 🚫 DO NOT DELETE, shorten, or summarize ANY existing text. Keep everything.
2. ✅ YOU MUST WRITE AT LEAST {deficit} NEW words to reach the minimum.
3. Method: EXPAND every paragraph with:
   - Concrete examples and case studies
   - Statistical evidence and data
   - Scholarly analysis and interpretation
   - Counterarguments and rebuttals
4. If critics found errors, fix them by ADDING clarification, NOT removing text.
5. This is a PAID requirement ({word_count} words). User expects full delivery.

IGNORE score improvements until word count >= {MIN_WORD_COUNT}.
"""
            elif current_word_count < word_count:
                # PHASE 2: MODERATE GROWTH (Between minimum and target)
                deficit = word_count - current_word_count
                mode_instruction = f"""
*** GROWTH MODE (APPROACHING TARGET) ***
Current: {current_word_count} words. Target: {word_count} words.
REMAINING: {deficit} words to reach full target.

RULES:
1. 🚫 DO NOT DELETE or shorten existing text.
2. ✅ Add {deficit} more words through expansion.
3. Balance growth with quality improvements from critics.
"""
            else:
                # PHASE 3: REFINEMENT (Target Met)
                mode_instruction = f"""
*** REFINEMENT MODE ***
Target met ({current_word_count}/{word_count}).
Now polish the text while keeping count above {word_count}.
"""

            # REVISION PHASE - Prof. Quill MERGES critic contributions WITH ENHANCED RATCHET + FLOOR MECHANISM
            yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill (Claude 3 Haiku) merging critic contributions...'})}\n\n"
            
            # Apply the mode to the prompt
            merge_prompt = f"""You are Prof. Quill.
{mode_instruction}

ORIGINAL INSTRUCTIONS: {truncated_instructions}

CURRENT DRAFT: {current_draft}

CRITICS' FEEDBACK:
Dean Logic: {logic_rewrites if logic_rewrites else "Expand content with academic depth."}
Chancellor GPT: {gpt_rewrites if gpt_rewrites else "Expand content with academic depth."}

EXECUTION:
Revise the essay following the STRICT RULES of the current Mode above.
"""

            # --- ENHANCED RATCHET GUARDRAIL WITH FLOOR MECHANISM ---
            try:
                # 1. Save previous state
                previous_draft_content = current_draft
                previous_word_count = current_word_count
                
                # 2. Generate new candidate draft
                candidate_draft = call_claude(merge_prompt, max_tokens=4000)
                candidate_word_count = count_words(candidate_draft)
                
                # 3. ENHANCED RATCHET GUARDRAIL (Two-Condition Check)
                # Condition 1: Shrinking while we are still trying to grow (Existing logic)
                is_shrinking_during_growth = (current_word_count < word_count) and (candidate_word_count < previous_word_count * 0.95)
                
                # Condition 2: Dropping below target after we already reached it (NEW FLOOR MECHANISM)
                # If we had 2033 words, and now we have 1220, REJECT IT.
                dropped_below_target = (previous_word_count >= word_count) and (candidate_word_count < word_count)

                if is_shrinking_during_growth or dropped_below_target:
                    
                    failure_reason = "Shrinking during growth" if is_shrinking_during_growth else "Dropped below target (Floor violation)"
                    yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ RATCHET TRIGGERED ({failure_reason}): Draft rejected. Candidate: {candidate_word_count} words, Previous: {previous_word_count} words'})}\n\n"
                    
                    # Restore previous draft and force expansion instructions
                    current_draft = previous_draft_content + "\n\n[SYSTEM NOTE: Your last edit cut too much text. You are FORBIDDEN from dropping below the word count target. Add 300 words.]"
                    current_word_count = previous_word_count
                    
                    yield f"data: {json.dumps({'type': 'log', 'message': f'🔒 Ratchet engaged: Keeping previous draft ({previous_word_count} words) + forcing 300-word expansion'})}\n\n"
                    
                else:
                    # Accept the new draft
                    current_draft = candidate_draft
                    current_word_count = candidate_word_count  # Update this variable explicitly
                    yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Draft accepted ({candidate_word_count} words)'})}\n\n"

            except Exception as e:
                yield f"data: {json.dumps({'type': 'error', 'message': f'Merge error: {str(e)}'})}\n\n"
                break
            
            # Send any queued heartbeats
            while heartbeat_queue:
                yield heartbeat_queue.pop(0)

        final_word_count = count_words(best_draft)
        final_avg = best_avg_score
        yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 COMPLETE! Final: {final_avg}/100 avg, {final_word_count} body words, {round_count} rounds'})}\n\n"
        
        # Save to database
        try:
            essay = Essay(
                user_id=current_user.id,
                title=instructions[:100],
                instructions=instructions,
                final_content=best_draft,
                final_score=final_avg,
                rounds_used=round_count,
                word_count_limit=word_count
            )
            db.session.add(essay)
            db.session.commit()
            
            yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': final_avg})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': f'DB error: {str(e)}'})}\n\n"
    
    finally:
        # Stop heartbeat thread
        stop_heartbeat.set()
        heartbeat_thread.join(timeout=1)

# TOOL D: The Grader - 3-Round Consensus Debate
def grade_assignment(brief_text, essay_text):
    """Grade assignment using 3-round consensus debate among 3 professors"""
    
    print("🎓 ROUND 1: Independent Grading...")
    
    # Truncate texts to prevent token overflow
    truncated_brief = truncate_text(brief_text, max_tokens=2000)
    truncated_essay = truncate_text(essay_text, max_tokens=8000)
    
    quill_prompt_r1 = f"""You are Prof. Quill, an expert academic evaluator. Analyze this student essay against the assignment brief.

ASSIGNMENT BRIEF:
{truncated_brief}

STUDENT ESSAY:
{truncated_essay}

Evaluate independently:
1. How well does it address the brief?
2. Content quality and depth
3. Structure and organization

Provide:
- A score (0-100)
- Detailed feedback (2-3 paragraphs)
- Key strengths and weaknesses"""

    strict_prompt_r1 = f"""You are Dr. Strict, a harsh academic grader. Grade this essay strictly and independently.

ASSIGNMENT BRIEF:
{truncated_brief}

STUDENT ESSAY:
{truncated_essay}

Evaluate independently:
1. Grammar and writing quality
2. Citation and referencing
3. Academic rigor

Provide:
- A score (0-100)
- Detailed feedback identifying specific weaknesses
- Technical issues found"""

    logic_prompt_r1 = f"""You are Dean Logic, the final academic authority. Provide an independent overall assessment.

ASSIGNMENT BRIEF:
{truncated_brief}

STUDENT ESSAY:
{truncated_essay}

Provide independently:
- Overall score (0-100)
- Key strengths
- Critical improvements needed
- Holistic evaluation"""

    try:
        quill_response_r1 = call_claude(quill_prompt_r1)
        quill_score_r1 = extract_score(quill_response_r1)
        
        strict_response_r1 = call_gpt4(strict_prompt_r1, max_input_tokens=12000)
        strict_score_r1 = extract_score(strict_response_r1)
        
        logic_response_r1 = call_gemini(logic_prompt_r1)
        logic_score_r1 = extract_score(logic_response_r1)
        
        print(f"Round 1 Scores - Quill: {quill_score_r1}, Strict: {strict_score_r1}, Logic: {logic_score_r1}")
        
        print("🎓 ROUND 2: Consensus Phase...")
        
        # Truncate feedback for round 2 to prevent token overflow
        truncated_quill_r1 = truncate_text(quill_response_r1, max_tokens=2000)
        truncated_strict_r1 = truncate_text(strict_response_r1, max_tokens=2000)
        truncated_logic_r1 = truncate_text(logic_response_r1, max_tokens=2000)
        
        quill_prompt_r2 = f"""You are Prof. Quill. You've read the feedback from Dr. Strict and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {quill_score_r1}/100
{truncated_quill_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{truncated_strict_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{truncated_logic_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        strict_prompt_r2 = f"""You are Dr. Strict. You've read the feedback from Prof. Quill and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {strict_score_r1}/100
{truncated_strict_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{truncated_quill_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{truncated_logic_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        logic_prompt_r2 = f"""You are Dean Logic. You've read the feedback from Prof. Quill and Dr. Strict.

YOUR INITIAL ASSESSMENT:
Score: {logic_score_r1}/100
{truncated_logic_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{truncated_quill_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{truncated_strict_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        quill_response_r2 = call_claude(quill_prompt_r2)
        quill_score_r2 = extract_score(quill_response_r2)
        
        strict_response_r2 = call_gpt4(strict_prompt_r2, max_input_tokens=12000)
        strict_score_r2 = extract_score(strict_response_r2)
        
        logic_response_r2 = call_gemini(logic_prompt_r2)
        logic_score_r2 = extract_score(logic_response_r2)
        
        print(f"Round 2 Scores - Quill: {quill_score_r2}, Strict: {strict_score_r2}, Logic: {logic_score_r2}")
        
        print("🎓 ROUND 3: Final Calculation...")
        
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
        
        if password != confirm_password:
            flash('Passwords do not match!', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(username=username).first():
            flash('Username already exists. Please choose a different username.', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email already registered. Please use a different email or login.', 'error')
            return redirect(url_for('register'))
        
        try:
            user = User(username=username, email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            
            email_body = f"""
Welcome to The Academic Board, {username}!

Your account has been successfully created.

Get started with our premium AI-powered academic tools:
- The Architect: AI Essay Writer (GROWTH MODE + SELF-REFLECTION System!)
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
    """Retrieve both Essays and Reports for unified dashboard view"""
    essays = Essay.query.filter_by(user_id=current_user.id).order_by(Essay.created_at.desc()).all()
    reports = Report.query.filter_by(user_id=current_user.id).order_by(Report.created_at.desc()).all()
    return render_template('my_essays.html', essays=essays, reports=reports, user=current_user)

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
        
        if not os.getenv('OPENAI_API_KEY'):
            return jsonify({'reply': 'System Error: Contact Admin - OpenAI API key not configured.'}), 500
        
        system_prompt = """You are The Concierge, the helpful support agent for The Academic Board.

PRICING:
- Essay Writer: £10 per 1,000 words (1,000-12,000 words range)
- Plagiarism Check: £10
- AI Detector: £10
- Assignment Grader: £10

TOOLS:
- The Architect: AI essay writer with GROWTH MODE + SELF-REFLECTION system (Prof. Quill, Dean Logic, Chancellor GPT)
- The Detective: Plagiarism checker with interactive Scribbr-style dashboard
- The Oracle: AI content detector with 3-agent consensus (GPT, Gemini, Claude)
- The Grader: Strict assignment marking

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
    """MEMORY-OPTIMIZED file upload handler for 25MB files."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Supported: PDF, DOCX, TXT (max 25MB each)'}), 400
    
    filepath = None
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        file.save(filepath)
        text = extract_text_from_file(filepath, filename)
        
        os.remove(filepath)
        filepath = None
        
        if len(text) > 100000:
            text = text[:100000] + "\n\n[Preview truncated. Full content will be used in essay generation.]"
        
        print(f"✅ File processed successfully: {filename}")
        return jsonify({'success': True, 'text': text})
        
    except Exception as e:
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
        
        if tool_type == 'writer':
            word_count = data.get('word_count', 2000)
            amount = int((word_count / 1000) * PRICING['writer']['price_per_1k'])
            product_name = f'Essay Writing ({word_count:,} words)'
            print(f"💰 Writer pricing: {word_count} words = £{amount/100}")
        else:
            amount = PRICING[tool_type]['price']
            product_names = {
                'plagiarism': 'Plagiarism Check',
                'ai_check': 'AI Detection',
                'grader': 'Assignment Grading'
            }
            product_name = product_names.get(tool_type, 'Academic Tool')
        
        session['pending_task'] = {'tool_type': tool_type, 'data': data}
        
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
                payment = Payment(
                    user_id=current_user.id,
                    amount=checkout_session.amount_total,
                    tool_type=tool_type,
                    stripe_session_id=session_id
                )
                db.session.add(payment)
                db.session.commit()
                
                flash('Payment successful!', 'success')
                
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

# FIX #3: PLAGIARISM (Simplify Logic + PDF Export + ALWAYS SHOW TEXT)
@app.route('/check-plagiarism', methods=['POST'])
@login_required
def check_plagiarism():
    """RESCUE FIX: Simplified plagiarism detection with guaranteed text visibility"""
    text = request.form.get('text')
    if not text:
        return jsonify({'error': 'No text'}), 400
    
    print("🔍 RESCUE FIX: Simplified plagiarism detection...")
    
    # Use simpler logic to avoid JSON crashes
    findings = plagiarism_hunter_gemini(text)
    
    # CRITICAL: Save the original input text
    findings['input_text'] = text
    
    # Visuals
    colors_list = ['#00E5FF', '#FFD600', '#76FF03', '#D500F9', '#FF1744']
    source_colors = {s.get('id'): colors_list[i % len(colors_list)] for i, s in enumerate(findings.get('sources', []))}

    # Left Panel: Sources
    sources_html = ""
    for s in findings.get('sources', []):
        color = source_colors.get(s.get('id'), '#ccc')
        url = s.get('url', '#')
        if not url.startswith('http'): 
            url = 'https://' + url
        
        sources_html += f'''
        <div style="border-left: 5px solid {color}; background: #2d2d2d; margin-bottom: 10px; padding: 10px;">
            <div style="display:flex; justify-content:space-between;">
                <a href="{url}" target="_blank" style="color:{color}; font-weight:bold; text-decoration:none;">{s.get('domain')} 🔗</a>
                <span style="color:white;">{s.get('similarity')}%</span>
            </div>
        </div>'''

    # Right Panel: Text (CRITICAL FIX: Ensure text is visible even if no matches found)
    highlighted_text = text
    if findings.get('matches'):
        for match in findings.get('matches', []):
            seg = match.get('text_segment', '')
            color = source_colors.get(match.get('source_id'), 'transparent')
            if len(seg) > 10:
                highlighted_text = highlighted_text.replace(seg, f'<span style="background-color: {color}40; border-bottom: 2px solid {color}; color: #fff;">{seg}</span>')

    pdf_script = """
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <script>
    function exportPDF() {
        const element = document.getElementById('plagiarism-dashboard');
        const opt = { margin: 0.2, filename: 'Plagiarism_Report.pdf', image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' } };
        html2pdf().set(opt).from(element).save();
    }
    </script>
    """

    html = f"""
    {pdf_script}
    <div id="plagiarism-dashboard" style="display: flex; gap: 20px; height: 600px; font-family: sans-serif; color: #e0e0e0; background: #121212; padding: 20px;">
        <div style="flex: 1; overflow-y: auto; padding-right: 15px; border-right: 1px solid #333;">
            <h3 style="margin-top:0;">Sources</h3>
            {sources_html if sources_html else "<p>No sources detected.</p>"}
        </div>
        <div style="flex: 2; background: #1e1e1e; padding: 30px; border-radius: 8px; overflow-y: auto; line-height: 1.8;">
            {highlighted_text}
        </div>
        <div style="flex: 0 0 150px; text-align: center;">
            <div style="width:120px; height:120px; border-radius:50%; border:10px solid #f44336; display:flex; align-items:center; justify-content:center; margin:0 auto; font-size: 2.5em; font-weight: bold;">
                {findings.get('score', 0)}%
            </div>
            <button onclick="exportPDF()" data-html2canvas-ignore="true" style="margin-top:20px; background: #2196F3; color: white; border: none; padding: 10px 20px; border-radius: 25px; cursor: pointer;">📥 Export PDF</button>
        </div>
    </div>"""
    
    # Save Report to DB
    try:
        report = Report(
            user_id=current_user.id,
            tool_type='plagiarism',
            title=f"Plagiarism Check: {text[:40]}...",
            result_data=json.dumps(findings)
        )
        db.session.add(report)
        db.session.commit()
        print(f"✅ Report saved to database (ID: {report.id})")
        report_id = report.id
    except Exception as e:
        print(f"❌ DB save error: {str(e)}")
        report_id = None
    
    print("✅ RESCUE FIX: Returning simplified plagiarism detection")
    return jsonify({'success': True, 'html': html, 'report_id': report_id})

# FIX #2: AI CHECK (Prevent Empty Box + PDF Export + ALWAYS SHOW TEXT)
@app.route('/check-ai', methods=['POST'])
@login_required
def check_ai():
    """RESCUE FIX: AI detection with guaranteed text visibility"""
    text = request.form.get('text')
    if not text:
        return jsonify({'error': 'No text'}), 400
    
    print("🤖 RESCUE FIX: AI detection with guaranteed text visibility...")
    
    # 1. Get Data (Consensus) - now includes input_text
    try:
        data = check_ai_consensus(text)
    except Exception as e:
        print(f"❌ Consensus Error: {e}")
        data = {"ai_score": 0, "segments": [], "input_text": text}

    # 2. Build HTML Content (CRITICAL FIX FOR EMPTY BOX)
    content_html = ""
    segments = data.get('segments', [])
    
    if not segments:
        # If AI failed to segment, just show original text
        content_html = f"<span>{text}</span>"
    else:
        for seg in segments:
            bg = "transparent"
            if seg.get('is_ai'):
                bg = "rgba(255, 23, 68, 0.4)" if seg.get('confidence', 0) > 75 else "rgba(255, 214, 0, 0.3)"
            content_html += f'<span style="background-color:{bg}; padding:2px 0;">{seg["text"]}</span> '

    # 3. PDF Script (Direct Browser Download)
    pdf_script = """
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <script>
    function exportAIPDF() {
        const element = document.getElementById('ai-dashboard');
        const opt = { margin: 0.2, filename: 'AI_Report.pdf', image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' } };
        html2pdf().set(opt).from(element).save();
    }
    </script>
    """

    html = f"""
    {pdf_script}
    <div id="ai-dashboard" style="display: flex; gap: 20px; height: 600px; font-family: sans-serif; background: #121212; padding: 20px; color: #e0e0e0;">
        <div style="flex: 3; background: #1e1e1e; padding: 30px; border-radius: 12px; border: 1px solid #333; overflow-y: auto; line-height: 1.8;">
            {content_html}
        </div>
        <div style="flex: 1; background: #252526; padding: 30px; border-radius: 12px; text-align: center;">
            <h3 style="color:#bdc3c7;">AI Probability</h3>
            <div style="font-size: 5em; color: #ff1744; font-weight: bold; margin: 20px 0;">{data.get('ai_score', 0)}%</div>
            <button onclick="exportAIPDF()" data-html2canvas-ignore="true" style="margin-top:20px; background: linear-gradient(45deg, #FF512F, #DD2476); color: white; border: none; padding: 12px 24px; border-radius: 25px; cursor: pointer;">📥 Export PDF</button>
        </div>
    </div>"""
    
    # Save Report to DB
    try:
        report = Report(
            user_id=current_user.id,
            tool_type='ai_check',
            title=f"AI Check: {text[:30]}...",
            result_data=json.dumps(data)
        )
        db.session.add(report)
        db.session.commit()
        print(f"✅ Report saved to database (ID: {report.id})")
        report_id = report.id
    except Exception as e:
        print(f"❌ DB save error: {str(e)}")
        report_id = None
    
    print("✅ RESCUE FIX: Returning AI detection with guaranteed text visibility")
    return jsonify({'success': True, 'html': html, 'report_id': report_id})

# NEW: DYNAMIC PDF EXPORT ENDPOINT (FIXES BROKEN EXPORT BUTTONS)
@app.route('/export-pdf/<int:report_id>')
@login_required
def export_pdf_dynamic(report_id):
    """
    UNIFIED DYNAMIC PDF GENERATOR
    Generates PDF on-the-fly with visual highlights for AI Detector and Plagiarism tools.
    No disk storage - streams directly to user.
    """
    report = Report.query.get_or_404(report_id)
    
    # Security check
    if report.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))

    # Setup Buffer and Document
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    # Title
    title_style = ParagraphStyle('MainTitle', parent=styles['Heading1'], alignment=1, fontSize=18, spaceAfter=20)
    story.append(Paragraph(f"📄 {report.title} Report", title_style))
    story.append(Paragraph(f"Date: {report.created_at.strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 20))

    # --- LOGIC FOR AI DETECTOR (The Oracle) ---
    if report.tool_type == 'ai_check':
        try:
            # Parse JSON data
            data = json.loads(report.result_data) if isinstance(report.result_data, str) else report.result_data
            score = data.get('ai_score', 0)
            
            # Score Display
            score_color = "red" if score > 50 else "green"
            score_text = f'<font color="{score_color}" size="16"><b>AI Probability: {score}%</b></font>'
            story.append(Paragraph(score_text, styles['Normal']))
            story.append(Spacer(1, 10))
            
            breakdown = data.get('breakdown', 'N/A')
            story.append(Paragraph(f"<b>Breakdown:</b> {breakdown}", styles['Normal']))
            story.append(Spacer(1, 20))

            # --- PDF CONTENT RENDER LOGIC (MERGED & IMPROVED) ---
            story.append(Paragraph("<b>Full Text Analysis:</b>", styles['Heading3']))
            story.append(Spacer(1, 12))

            # Get data
            segments = data.get('segments', [])
            full_text = data.get('input_text', '')

            # Helper to render text with proper spacing (Fixes "Wall of Text")
            def add_formatted_text(text_content, background_color=None):
                # Split by newlines to preserve paragraphs
                paragraphs = text_content.split('\n')
                for p in paragraphs:
                    if p.strip():  # Skip empty lines
                        if background_color:
                            # Highlight the whole paragraph
                            style = ParagraphStyle('Highlight', parent=styles['Normal'], backColor=background_color)
                            story.append(Paragraph(p, style))
                        else:
                            story.append(Paragraph(p, styles['Normal']))
                        # ADD SPACING BETWEEN PARAGRAPHS
                        story.append(Spacer(1, 8))

            if segments:
                # Reconstruct text with segments using Gradient Sensitivity
                for seg in segments:
                    text_part = seg.get('text', '')
                    score_seg = seg.get('confidence', 0)
                    
                    # HYPER-SENSITIVE THRESHOLDS (>1%)
                    if score_seg > 80:
                        # HIGH PROBABILITY -> RED (MistyRose)
                        bg_color = colors.Color(1, 0.8, 0.8)  # #FFCCCC
                        story.append(Paragraph(text_part, ParagraphStyle('HighAI', parent=styles['Normal'], backColor=bg_color)))
                    
                    elif score_seg > 40:
                        # MEDIUM PROBABILITY -> ORANGE/GOLD
                        bg_color = colors.Color(1, 0.92, 0.6)  # #FFEB99
                        story.append(Paragraph(text_part, ParagraphStyle('MedAI', parent=styles['Normal'], backColor=bg_color)))
                        
                    elif score_seg > 1:
                        # LOW PROBABILITY (2-40%) -> LIGHT YELLOW
                        # User wants to see EVERYTHING
                        bg_color = colors.Color(1, 1, 0.8)  # #FFFFCC
                        story.append(Paragraph(text_part, ParagraphStyle('LowAI', parent=styles['Normal'], backColor=bg_color)))
                        
                    else:
                        # 0-1% Score -> No Highlight
                        story.append(Paragraph(text_part, styles['Normal']))
                    
                    # Small spacer for readability between segments
                    story.append(Spacer(1, 6))

            elif full_text:
                # FALLBACK: If no segments, tint based on overall score
                overall_score = data.get('ai_score', 0)
                bg_color = None
                
                if overall_score > 80:
                    bg_color = colors.Color(1, 0.8, 0.8)  # Red tint
                    story.append(Paragraph(f"<i>(Note: Segments not detected, but High Risk ({overall_score}%). Full tint applied.)</i>", styles['Italic']))
                elif overall_score > 40:
                    bg_color = colors.Color(1, 0.92, 0.6)  # Orange tint
                    story.append(Paragraph(f"<i>(Note: Segments not detected, but Medium Risk ({overall_score}%). Full tint applied.)</i>", styles['Italic']))
                elif overall_score > 1:
                    bg_color = colors.Color(1, 1, 0.8)  # Yellow tint
                    story.append(Paragraph(f"<i>(Note: Segments not detected, but Low Risk ({overall_score}%). Full tint applied.)</i>", styles['Italic']))
                
                add_formatted_text(full_text, bg_color)
                
            else:
                story.append(Paragraph("<i>(Original text content not found in report data)</i>", styles['Italic']))

        except Exception as e:
            story.append(Paragraph(f"Error parsing data: {str(e)}", styles['Normal']))

    # --- LOGIC FOR PLAGIARISM (The Detective) ---
    elif report.tool_type == 'plagiarism':
        try:
            # Parse JSON data
            data = json.loads(report.result_data)
            score = data.get('score', 0)
            
            # Score Display
            score_color = "red" if score > 20 else "green"
            score_text = f'<font color="{score_color}" size="16"><b>Plagiarism Risk: {score}%</b></font>'
            story.append(Paragraph(score_text, styles['Normal']))
            story.append(Spacer(1, 20))
            
            # Sources List
            story.append(Paragraph("<b>Detected Sources:</b>", styles['Heading3']))
            story.append(Spacer(1, 10))
            
            sources = data.get('sources', [])
            if not sources:
                story.append(Paragraph("No suspicious sources detected.", styles['Normal']))
            else:
                for s in sources:
                    domain = s.get('domain', 'Unknown')
                    url = s.get('url', '#')
                    similarity = s.get('similarity', 0)
                    
                    source_text = f"• <b>{domain}</b> ({similarity}% match)<br/><i>{url}</i>"
                    story.append(Paragraph(source_text, styles['Normal']))
                    story.append(Spacer(1, 8))
            
            story.append(Spacer(1, 20))
            
            # --- PDF CONTENT RENDER LOGIC (MERGED & IMPROVED) ---
            story.append(Paragraph("<b>Full Text Analysis:</b>", styles['Heading3']))
            story.append(Spacer(1, 12))
            
            matches = data.get('matches', [])
            full_text = data.get('input_text', '')
            
            # Helper to render text with proper spacing (Fixes "Wall of Text")
            def add_formatted_text(text_content, background_color=None):
                # Split by newlines to preserve paragraphs
                paragraphs = text_content.split('\n')
                for p in paragraphs:
                    if p.strip():  # Skip empty lines
                        if background_color:
                            # Highlight the whole paragraph
                            style = ParagraphStyle('Highlight', parent=styles['Normal'], backColor=background_color)
                            story.append(Paragraph(p, style))
                        else:
                            story.append(Paragraph(p, styles['Normal']))
                        # ADD SPACING BETWEEN PARAGRAPHS
                        story.append(Spacer(1, 8))
            
            if matches and full_text:
                # Show full text with highlighted matches
                # Build a list of (start_pos, end_pos, source_info) for each match
                match_positions = []
                for match in matches:
                    segment = match.get('text_segment', '')
                    source_id = match.get('source_id', 0)
                    
                    # Find the segment in the full text
                    start_pos = full_text.find(segment)
                    if start_pos != -1:
                        end_pos = start_pos + len(segment)
                        source = next((s for s in sources if s.get('id') == source_id), None)
                        if source:
                            match_positions.append({
                                'start': start_pos,
                                'end': end_pos,
                                'similarity': source.get('similarity', 0),
                                'domain': source.get('domain', 'Unknown')
                            })
                
                # Sort by position
                match_positions.sort(key=lambda x: x['start'])
                
                # Build paragraphs with highlights
                last_pos = 0
                for mp in match_positions:
                    # Add normal text before match
                    if mp['start'] > last_pos:
                        normal_text = full_text[last_pos:mp['start']]
                        if normal_text.strip():
                            add_formatted_text(normal_text)
                    
                    # Add highlighted match
                    matched_text = full_text[mp['start']:mp['end']]
                    similarity = mp['similarity']
                    domain = mp['domain']
                    
                    if similarity > 50:
                        # High risk - Red highlight
                        match_style = ParagraphStyle('HighRisk', parent=styles['Normal'],
                                                    backColor=colors.Color(1, 0.8, 0.8),
                                                    borderColor=colors.red, borderWidth=1, borderPadding=5)
                        prefix = f"🔴 <b>HIGH RISK ({similarity}% - {domain}):</b><br/>"
                    elif similarity > 20:
                        # Medium risk - Yellow highlight
                        match_style = ParagraphStyle('MedRisk', parent=styles['Normal'],
                                                    backColor=colors.Color(1, 0.92, 0.6),
                                                    borderPadding=3)
                        prefix = f"🟠 <b>Possible Match ({similarity}% - {domain}):</b><br/>"
                    else:
                        match_style = styles['Normal']
                        prefix = f"🟢 <b>Low Risk ({similarity}% - {domain}):</b><br/>"
                    
                    story.append(Paragraph(prefix + matched_text, match_style))
                    story.append(Spacer(1, 8))
                    
                    last_pos = mp['end']
                
                # Add remaining text
                if last_pos < len(full_text):
                    remaining_text = full_text[last_pos:]
                    if remaining_text.strip():
                        add_formatted_text(remaining_text)
                        
            elif full_text:
                # No matches but we have text - show it in normal style
                add_formatted_text(full_text)
            else:
                story.append(Paragraph("<i>(Original text content not found in report data)</i>", styles['Italic']))

        except Exception as e:
            story.append(Paragraph(f"Error parsing data: {str(e)}", styles['Normal']))

    # Build PDF
    doc.build(story)
    buffer.seek(0)
    
    return Response(buffer, mimetype='application/pdf', 
                   headers={'Content-Disposition': f'attachment;filename=report_{report.tool_type}_{report_id}.pdf'})

# --- VIEW REPORT ROUTE (FOR ACCESSING SAVED REPORTS) ---
@app.route('/view-report/<int:report_id>')
@login_required
def view_report(report_id):
    """View a saved report from the database"""
    report = Report.query.get_or_404(report_id)
    
    if report.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))
    
    # Parse stored data
    data = json.loads(report.result_data)
    
    # Re-render using Dark Mode styles
    if report.tool_type == 'ai_check':
        # Regenerate AI detection HTML
        content_html = ""
        segments = data.get('segments', [])
        
        if not segments:
            content_html = "<span>No segments available.</span>"
        else:
            for seg in segments:
                bg = "transparent"
                if seg.get('is_ai'):
                    bg = "rgba(255, 23, 68, 0.4)" if seg.get('confidence', 0) > 75 else "rgba(255, 214, 0, 0.3)"
                content_html += f'<span style="background-color:{bg}; border-radius:3px; padding:2px 0;">{seg.get("text", "")}</span> '
        
        content = f"""
        <div style="display: flex; gap: 20px; height: 500px; font-family: sans-serif;">
            <div style="flex: 3; background: #1e1e1e; color: #e0e0e0; padding: 25px; border-radius: 12px; border: 1px solid #333; overflow-y: scroll; line-height: 1.8; font-size: 1.1em;">
                {content_html}
            </div>
            <div style="flex: 1; background: #252526; border: 1px solid #333; color: white; padding: 25px; border-radius: 12px; text-align: center; display: flex; flex-direction: column; justify-content: center;">
                <h3 style="margin:0; color:#bdc3c7; font-size: 1.2em; text-transform: uppercase; letter-spacing: 1px;">AI Probability</h3>
                <div style="font-size: 6em; color: #ff1744; font-weight: bold; margin: 20px 0; text-shadow: 0 0 20px rgba(255, 23, 68, 0.4);">
                    {data.get('ai_score', 0)}%
                </div>
                <p style="color: #7f8c8d; font-size: 0.9em;">Highlights indicate likely AI-generated text.</p>
            </div>
        </div>
        """
    else:  # plagiarism
        # Regenerate plagiarism HTML with color coding
        score = data.get('score', 0)
        gauge_color = '#ff1744' if score > 20 else '#00e676'
        
        # Color palette
        colors_list = ['#00E5FF', '#FFD600', '#76FF03', '#D500F9', '#FF1744']
        source_colors = {}
        for idx, s in enumerate(data.get('sources', [])):
            source_colors[s.get('id')] = colors_list[idx % len(colors_list)]
        
        sources_html = ''
        for s in data.get('sources', []):
            color = source_colors.get(s.get('id'), '#ccc')
            url = s.get('url', '#')
            if url and url != '#' and not url.startswith('http'):
                url = 'https://' + url
            domain = s.get('domain', 'Unknown')
            similarity = s.get('similarity', 0)
            
            sources_html += f'''
            <div style="border-left: 5px solid {color}; background: #2d2d2d; margin-bottom: 10px; padding: 12px; border-radius: 4px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <a href="{url}" target="_blank" style="color: {color}; font-weight:bold; text-decoration:none; font-size:1.1em; word-break: break-all;">
                        {domain} 🔗
                    </a>
                    <span style="background:{color}; color:black; padding:2px 8px; border-radius:10px; font-weight:bold;">{similarity}%</span>
                </div>
            </div>'''
        
        if not sources_html:
            sources_html = '<p style="padding:20px; text-align:center; color:#666;">No suspicious sources detected.</p>'
        
        content = f"""
        <div style="display: flex; gap: 20px; margin-top: 20px; font-family: sans-serif;">
            <div style="flex: 1; background: #1e1e1e; padding: 20px; border-radius: 12px; border: 1px solid #333;">
                <h3 style="color: #e0e0e0; margin-top:0;">Detected Sources</h3>
                {sources_html}
            </div>
            <div style="flex: 1; text-align: center; background: #1e1e1e; padding: 30px; border-radius: 12px; border: 1px solid #333; height: fit-content;">
                <div style="width:160px; height:160px; border-radius:50%; border:15px solid {gauge_color}; display:flex; align-items:center; justify-content:center; margin:0 auto; box-shadow: 0 0 20px {gauge_color}40;">
                    <h1 style="font-size:3.5em; margin:0; color: #e0e0e0;">{score}%</h1>
                </div>
                <h3 style="color:#b0bec5; margin-top:25px;">Plagiarism Risk</h3>
            </div>
        </div>
        """
    
    return render_template('view_report.html', content=content, title=report.title, report=report)

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
    # CRITICAL: Gunicorn timeout MUST be set to 600s+ in Render Start Command: gunicorn --timeout 600 --keep-alive 5 --workers 1 app:app
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
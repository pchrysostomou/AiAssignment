# AI Assignment Architect - PRO SaaS

A commercial-grade AI writing tool that uses a multi-model approach to create high-quality essays through an intelligent refinement loop.

## 🚀 Features

- **Elite Squad AI Models**: Claude 3.5 Sonnet (Writer), GPT-4o (Examiner), Gemini 1.5 Pro (Reviewer)
- **Intelligent Refinement Loop**: Up to 10 rounds with smart early stopping
- **Cost Optimization**: Context pruning from round 4 onwards
- **Real-time Streaming**: Live progress updates via Server-Sent Events
- **Quality Scoring**: Automatic 0-100 grading system
- **Smart Break Conditions**:
  - Success: Score ≥ 90
  - Stagnation: <1.5 point improvement over 2 rounds
  - Limit: Maximum 10 rounds
- **Word Export**: Professional document generation with python-docx
- **User Management**: Secure authentication with Flask-Login

## 📋 Prerequisites

- Python 3.8+
- Windows 11 / VS Code (as specified)
- API Keys for:
  - Anthropic (Claude)
  - OpenAI (GPT-4o)
  - Google AI (Gemini)

## 🛠️ Installation

1. **Clone or create the project directory**
```bash
mkdir ai-assignment-architect
cd ai-assignment-architect
```

2. **Create a virtual environment**
```bash
python -m venv venv
venv\Scripts\activate  # Windows
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Configure environment variables**
```bash
# Copy the template
copy .env.template .env

# Edit .env and add your API keys
notepad .env
```

Required API keys in `.env`:
```
ANTHROPIC_API_KEY=sk-ant-...
OPENAI_API_KEY=sk-...
GOOGLE_API_KEY=AI...
SECRET_KEY=your-random-secret-key
```

5. **Initialize the database**
```bash
python app.py
```

This will create:
- SQLite database (`ai_architect.db`)
- Demo user account (username: `demo`, password: `demo123`)

## 🎯 Usage

### Starting the Application

```bash
python app.py
```

The application will be available at: `http://localhost:5000`

### Login

Use the demo account or register a new one:
- Username: `demo`
- Password: `demo123`

### Creating an Essay

1. Enter your essay instructions in the text area
2. Click "Generate Essay (£20)"
3. Watch the real-time progress logs
4. Download the final Word document

### Understanding the Logs

The streaming logs show:
- Initial draft creation
- Round-by-round scores
- Examiner feedback
- Reviewer critiques
- Early stopping decisions
- Context pruning activation (Round 4+)

## 🧠 How It Works

### The Logic Loop

```
1. Writer (Claude 3.5 Sonnet) creates Draft 1
2. Loop begins (Max 10 rounds):
   a. Examiner (GPT-4o) grades (0-100)
   b. Check break conditions:
      - SUCCESS: Score ≥ 90 → STOP
      - STAGNATION: <1.5 improvement → STOP
      - LIMIT: Round 10 → STOP
   c. Reviewer (Gemini 1.5 Pro) provides critique
   d. Writer refines based on feedback
3. Save best version to database
```

### Context Optimization

**Rounds 1-3**: Full conversation history sent to AI
**Rounds 4+**: Pruned context (saves API costs):
- Original instructions
- Current draft only
- Previous round's feedback only

This reduces token usage by ~60% while maintaining quality.

## 📊 API Cost Estimates

Per essay generation (approximate):

- **Best case** (2-3 rounds, early success):
  - Claude: ~8K tokens ($0.10)
  - GPT-4o: ~4K tokens ($0.08)
  - Gemini: ~4K tokens ($0.01)
  - **Total: ~$0.19**

- **Average case** (5-7 rounds):
  - Claude: ~20K tokens ($0.25)
  - GPT-4o: ~10K tokens ($0.20)
  - Gemini: ~10K tokens ($0.03)
  - **Total: ~$0.48**

- **Worst case** (10 rounds, no early stop):
  - Claude: ~35K tokens ($0.44)
  - GPT-4o: ~18K tokens ($0.36)
  - Gemini: ~18K tokens ($0.05)
  - **Total: ~$0.85**

Context pruning saves approximately **40-60% on API costs** for rounds 4+.

## 🔧 Configuration

Edit these values in `.env` to customize behavior:

```env
MAX_ROUNDS=10                    # Maximum refinement rounds
SUCCESS_THRESHOLD=90             # Score to trigger early stop
STAGNATION_THRESHOLD=1.5         # Min improvement to continue
HISTORY_PRUNING_START=4          # Round to start pruning context
```

## 🐛 Error Handling

The application includes:
- Automatic retry logic (3 attempts) for API failures
- Graceful degradation on errors
- Detailed error logging in the UI
- Database transaction rollback on failures

## 📁 Project Structure

```
ai-assignment-architect/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── .env.template         # Environment variables template
├── .env                  # Your actual config (not in git)
├── ai_architect.db       # SQLite database (auto-created)
├── templates/
│   ├── login.html        # Login/Register page
│   └── dashboard.html    # Main dashboard
└── README.md            # This file
```

## 🔒 Security Notes

- Change `SECRET_KEY` in production
- Use environment variables for API keys
- Never commit `.env` to version control
- Consider using PostgreSQL for production
- Add rate limiting for production use
- Implement proper payment processing for £20/use

## 🚀 Production Deployment

For production deployment:

1. Use a production WSGI server (Gunicorn/uWSGI)
2. Set up proper database (PostgreSQL)
3. Configure reverse proxy (Nginx)
4. Enable HTTPS
5. Implement payment gateway
6. Add monitoring and logging
7. Set up backup system

Example with Gunicorn:
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## 📝 License

Commercial use - £20 per essay generation

## 🤝 Support

For issues or questions, contact support or check the logs in the dashboard.

## 🎓 Academic Integrity

This tool is designed for educational assistance. Users are responsible for ensuring their use complies with academic integrity policies.
# 🚨 CRITICAL: Gunicorn Timeout Configuration Required

## Problem
The application is experiencing "Failed to fetch" errors during:
- **Stripe Payment Redirect**: CSP blocking eval() and script-src
- **AI Essay Generation**: Timeout during "Scanning Academic Database" and "Final Polish" phases

## Root Cause
**Gunicorn's default 30-second timeout is killing long-running AI operations** that take 2-5 minutes to complete (5 rounds of Claude drafting + GPT-4 grading + Gemini citation verification).

## ✅ CSP Already Fixed in app.py
The Content Security Policy in `app.py` (lines 54-63) already includes:
- ✅ `'unsafe-eval'` in script-src (required for Stripe.js)
- ✅ `'unsafe-inline'` in script-src (required for inline scripts)
- ✅ `https://js.stripe.com` whitelisted in script-src and frame-src
- ✅ Your domain in connect-src for API calls
- ✅ `https://api.stripe.com` in connect-src for Stripe API
- ✅ `https://www.googleapis.com` in connect-src for Google Search

**The CSP is NOT the problem.** The timeout is.

---

## 🔧 MANDATORY FIX: Increase Gunicorn Timeout to 300s

### Option 1: Render Dashboard (Recommended)

1. **Login to Render Dashboard**
   - Go to https://dashboard.render.com
   - Select your service: `aiassignment-x631`

2. **Navigate to Settings**
   - Click "Settings" tab in the left sidebar

3. **Update Start Command**
   - Scroll to "Build & Deploy" section
   - Find "Start Command" field
   - **Replace the current command with:**
   ```bash
   gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 2 --keep-alive 5 app:app
   ```

4. **Explanation of Flags:**
   - `--timeout 300`: 5 minutes (300 seconds) for long AI operations
   - `--workers 2`: 2 worker processes for concurrent requests
   - `--keep-alive 5`: Keep connections alive for 5 seconds

5. **Save and Deploy**
   - Click "Save Changes"
   - Render will automatically redeploy (takes 2-3 minutes)

---

### Option 2: Environment Variable (Alternative)

If you prefer using environment variables:

1. Go to Render Dashboard → Your Service → Environment
2. Add new environment variable:
   - **Key**: `GUNICORN_CMD_ARGS`
   - **Value**: `--timeout 300 --workers 2 --keep-alive 5`
3. Click "Save Changes"
4. Render will redeploy automatically

---

## 📊 Why 300 Seconds?

The AI essay generation process involves:

| Phase | Time | Details |
|-------|------|---------|
| Google Custom Search | 10-15s | Research phase |
| Claude Initial Draft | 30-60s | Prof. Quill writing |
| **Round 1-5 (each):** | | |
| - GPT-4 Grading | 20-30s | Dr. Strict evaluation |
| - Gemini Citation Check | 15-20s | Dean Logic verification |
| - Claude Revision | 30-60s | Prof. Quill rewriting |
| Database Save | 2-5s | Final storage |

**Total: 2-5 minutes per essay**

With the default 30s timeout, the connection drops during Round 2-3, causing "Failed to fetch" errors.

---

## 🧪 Testing After Deployment

### 1. Test Stripe Payment
```
1. Open browser console (F12)
2. Navigate to The Architect tool
3. Fill in essay details (2,000 words)
4. Click "Pay & Generate Essay"
5. Expected: Stripe checkout opens (no CSP errors)
6. Enter promo code: MAKIS20
7. Complete payment
```

### 2. Test AI Generation
```
1. After payment, essay generation starts
2. Watch console for real-time logs:
   - "🔍 Researching topic via Google Custom Search..."
   - "✅ Found 5 relevant sources"
   - "✍️ Prof. Quill drafting (Spartan precision)..."
   - "🎯 Dr. Strict grading..."
   - "📊 Dr. Strict's verdict: 85/100"
   - "🔍 Dean Logic verifying citations..."
   - "✍️ Prof. Quill revising (Spartan mode)..."
3. Process completes in 2-5 minutes
4. Download button appears
```

### 3. Verify No Errors
```
- Console: NO CSP violations
- Console: NO "Failed to fetch" errors
- Network tab: All requests return 200 OK
- Render logs: No timeout errors
```

---

## 🚨 If Issues Persist

### Check Render Logs
1. Go to Render Dashboard → Your Service → Logs
2. Look for these error patterns:
   ```
   [CRITICAL] WORKER TIMEOUT (pid:12345)
   Worker with pid 12345 was terminated due to signal 9
   ```
3. If you see these, the timeout setting didn't apply correctly

### Verify Start Command
1. In Render Settings, confirm Start Command shows:
   ```
   gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 2 --keep-alive 5 app:app
   ```
2. If not, update it and save

### Check Environment Variables
1. Verify these are set in Render → Environment:
   - `STRIPE_PUBLISHABLE_KEY` = `pk_live_...`
   - `STRIPE_SECRET_KEY` = `sk_live_...`
   - `ANTHROPIC_API_KEY` = (your key)
   - `OPENAI_API_KEY` = (your key)
   - `GOOGLE_API_KEY` = (your key)
   - `GOOGLE_CSE_ID` = `50f552d88c3e14772`

---

## 📝 Summary

**The CSP is already correctly configured in app.py.** The issue is purely a Gunicorn timeout problem.

**Action Required:**
1. Update Gunicorn start command in Render Settings to include `--timeout 300`
2. Redeploy
3. Test payment and AI generation

**Expected Result:**
- Stripe payment works (CSP allows eval())
- AI generation completes without timeout (300s allows 5-minute operations)
- No "Failed to fetch" errors

---

## 🆘 Support

If you still encounter issues after following these steps:
1. Check Render logs for specific error messages
2. Verify the Start Command was saved correctly
3. Ensure all environment variables are set
4. Contact Render support if the timeout setting doesn't apply
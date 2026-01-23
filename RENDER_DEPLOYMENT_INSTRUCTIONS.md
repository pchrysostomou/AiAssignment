# Render Deployment Instructions for Gunicorn Timeout

## Critical Configuration Required

To prevent "Failed to fetch" errors during long-running AI essay generation (which takes 2-5 minutes), you MUST increase the Gunicorn timeout in your Render service settings.

### Step-by-Step Instructions:

1. **Go to Render Dashboard**
   - Navigate to https://dashboard.render.com
   - Select your service: `aiassignment-x631`

2. **Update Environment Variables**
   - Click on "Environment" tab
   - Verify these keys are set:
     - `STRIPE_PUBLISHABLE_KEY` = `pk_live_...` (your Stripe publishable key)
     - `STRIPE_SECRET_KEY` = `sk_live_...` (your Stripe secret key)
     - `ANTHROPIC_API_KEY` = (your Anthropic API key)
     - `OPENAI_API_KEY` = (your OpenAI API key)
     - `GOOGLE_API_KEY` = (your Google API key)
     - `GOOGLE_CSE_ID` = `50f552d88c3e14772`

3. **Configure Gunicorn Timeout**
   - Click on "Settings" tab
   - Scroll to "Build & Deploy" section
   - Find "Start Command" field
   - Replace the default command with:
     ```
     gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 2 app:app
     ```
   - This sets:
     - `--timeout 300`: 5 minutes (300 seconds) for long AI operations
     - `--workers 2`: 2 worker processes for handling concurrent requests

4. **Save and Deploy**
   - Click "Save Changes"
   - Render will automatically redeploy your service
   - Wait for deployment to complete (usually 2-3 minutes)

## Why This Is Necessary

The AI essay generation process involves:
1. Google Custom Search (10-15 seconds)
2. Claude drafting (30-60 seconds)
3. GPT-4o grading (20-30 seconds per round × 5 rounds)
4. Gemini citation verification (15-20 seconds per round × 5 rounds)
5. Claude revision (30-60 seconds per round × 5 rounds)

**Total time: 2-5 minutes**

The default Gunicorn timeout is 30 seconds, which causes the connection to drop before the essay is complete, resulting in "Failed to fetch" errors.

## CSP Configuration

The updated `app.py` now includes a properly configured Content Security Policy that:
- ✅ Allows `'unsafe-eval'` for Stripe.js payment processing
- ✅ Allows `https://js.stripe.com` for Stripe scripts and iframes
- ✅ Allows your domain for API calls (`connect-src`)
- ✅ Maintains reasonable security for other resources

## Testing After Deployment

1. **Test Stripe Payment:**
   - Navigate to The Architect tool
   - Fill in essay details (2,000 words)
   - Click "Pay & Generate Essay"
   - Verify Stripe checkout page opens
   - Enter promo code MAKIS20
   - Complete payment

2. **Test AI Generation:**
   - After payment, essay generation should start automatically
   - You should see real-time logs in the console
   - Process should complete in 2-5 minutes without "Failed to fetch" errors

3. **Check Console for Errors:**
   - Open browser console (F12)
   - Verify no CSP violations
   - Verify no "Failed to fetch" errors

## Troubleshooting

If you still see "Failed to fetch" errors after deployment:

1. **Check Render Logs:**
   - Go to Render Dashboard → Your Service → Logs
   - Look for timeout errors or worker crashes

2. **Verify Gunicorn Command:**
   - In Render Settings, confirm Start Command shows `--timeout 300`

3. **Check Environment Variables:**
   - Ensure all API keys are set correctly
   - Verify STRIPE_PUBLISHABLE_KEY is present

4. **Browser Console:**
   - Check for CSP violations (should be none)
   - Check Network tab for failed requests
   - Look for specific error messages

## Support

If issues persist after following these instructions, contact Render support or check the application logs for specific error messages.
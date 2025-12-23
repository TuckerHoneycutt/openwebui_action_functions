# Troubleshooting Guide

## Where to Find Logs

### 1. OpenWebUI Server Logs

**Docker Installation:**
```bash
# View logs
docker logs <container_name> -f

# Or if using docker-compose
docker-compose logs -f open-webui
```

**Local Installation:**
- Check the terminal/console where you started OpenWebUI
- Logs are typically printed to stdout/stderr
- Look for lines starting with `[DOC_FORMATTER]` for our function's debug output

**Log File Locations:**
- `~/.open-webui/logs/` (if logging to file is enabled)
- Check OpenWebUI configuration for custom log paths

### 2. Browser Console (Client-Side)

1. Open your browser's Developer Tools:
   - **Chrome/Edge**: Press `F12` or `Ctrl+Shift+I` (Windows/Linux) / `Cmd+Option+I` (Mac)
   - **Firefox**: Press `F12` or `Ctrl+Shift+K` (Windows/Linux) / `Cmd+Option+K` (Mac)
   - **Safari**: Enable Developer menu first, then `Cmd+Option+I`

2. Check the **Console** tab for JavaScript errors

3. Check the **Network** tab:
   - Click the action button
   - Look for API calls to `/api/v1/actions/` or similar
   - Check the response for errors

### 3. OpenWebUI Function Status

1. Go to **Settings** → **Functions** (or **Workspace** → **Functions**)
2. Verify that "Document Style Formatter" appears in the list
3. Check if there are any error indicators
4. Ensure the function is enabled/activated

## Common Issues

### Button Doesn't Appear

1. **Function Not Registered:**
   - Check that `main.py` is in the correct directory: `~/.open-webui/functions/`
   - Restart OpenWebUI completely
   - Check server logs for import errors

2. **Function Not Assigned to Model:**
   - Go to **Workspace** → **Models**
   - Select your model
   - Ensure the action function is assigned/enabled

3. **Dependencies Missing:**
   - Check server logs for import errors
   - Verify all dependencies are installed: `pip install python-docx PyMuPDF pdf2docx pydantic`

### Button Clicks But Nothing Happens

1. **Check Browser Console:**
   - Look for JavaScript errors
   - Check Network tab for failed API calls

2. **Check Server Logs:**
   - Look for `[DOC_FORMATTER]` debug messages
   - Check for Python exceptions/tracebacks

3. **Verify Action Method is Called:**
   - The logs should show: `[DOC_FORMATTER] Action called with body keys: ...`
   - If you don't see this, the action isn't being triggered

### GUI Doesn't Display

1. **HTML Return Format:**
   - OpenWebUI may expect HTML in different formats
   - Check server logs for the return value
   - The function returns HTML in `html`, `result`, and `content` keys

2. **Browser Security:**
   - Some browsers block inline scripts
   - Check browser console for CSP (Content Security Policy) errors

3. **JavaScript Errors:**
   - Check browser console for JavaScript errors
   - The GUI uses modern JavaScript - ensure browser is up to date

## Debug Steps

1. **Enable Verbose Logging:**
   - The function already includes debug prints to stderr
   - Look for `[DOC_FORMATTER]` prefix in logs

2. **Test Action Directly:**
   ```python
   # In Python shell or test script
   from main import Action
   action = Action()
   result = await action.action({"test": "data"})
   print(result)
   ```

3. **Check Function Structure:**
   - Verify `Action` class exists
   - Verify `Valves` nested class exists
   - Verify `action` method signature matches expected format

4. **Verify Dependencies:**
   ```bash
   python -c "from docx import Document; print('python-docx OK')"
   python -c "from pdf2docx import Converter; print('pdf2docx OK')"
   python -c "from pydantic import BaseModel; print('pydantic OK')"
   ```

## Getting Help

If issues persist:

1. **Collect Information:**
   - Server logs (especially lines with `[DOC_FORMATTER]`)
   - Browser console errors
   - OpenWebUI version
   - Python version
   - Operating system

2. **Check OpenWebUI Documentation:**
   - https://docs.openwebui.com/features/plugin/functions/action/

3. **Community Support:**
   - OpenWebUI GitHub Discussions
   - OpenWebUI Discord

# CPM Criteria ID Extractor

A Streamlit application for extracting Criteria IDs from CPM (Care Plan Management) documents.

## Features

- Upload and process CPM documents (PDF, DOCX, DOCM formats)
- Extract Criteria IDs based on checkbox states
- Save CPMs to a temporary library
- Export extracted IDs to Excel
- Debug mode for troubleshooting extractions

## Cloud Deployment to Streamlit Community Cloud

### Prerequisites

1. A GitHub account
2. Your application files ready to upload

### Step-by-Step Deployment Instructions

#### Step 1: Create a GitHub Repository

1. Go to [GitHub](https://github.com) and log in
2. Click the "+" icon in the top-right corner and select "New repository"
3. Name your repository (e.g., `cpm-extractor`)
4. Choose "Public" (required for free Streamlit Community Cloud)
5. **Do NOT** initialize with README, .gitignore, or license (we already have these files)
6. Click "Create repository"

#### Step 2: Upload Your Files to GitHub

You have two options:

**Option A: Upload via Web Interface (Easier)**

1. On your new repository page, click "uploading an existing file"
2. Drag and drop these files:
   - `cpm_app.py`
   - `requirements.txt`
   - `.gitignore`
   - `README.md` (this file)
3. Scroll down and click "Commit changes"

**Option B: Use Git Command Line**

```bash
# In your project folder, run:
git init
git add cpm_app.py requirements.txt .gitignore README.md
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

#### Step 3: Deploy to Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click "New app"
4. Fill in the deployment form:
   - **Repository:** Select your GitHub repository
   - **Branch:** main
   - **Main file path:** cpm_app.py
5. Click "Deploy!"

#### Step 4: Wait for Deployment

- The deployment process takes 2-5 minutes
- You'll see build logs in real-time
- Once complete, you'll get a URL like: `https://YOUR-APP-NAME.streamlit.app`

#### Step 5: Share Your App

- Copy the URL and share it with your colleague
- The app will be publicly accessible (no login required)
- You can manage your apps at [share.streamlit.io](https://share.streamlit.io)

## Important Notes

### File Storage Limitations

⚠️ **The CPM Library is temporary!** Files saved to the cloud instance will be deleted when:
- The app restarts (after 7 days of inactivity)
- You redeploy the app
- Streamlit performs maintenance

**Recommendation:** Always download your extracted IDs immediately after processing.

### For Persistent Storage (Advanced)

If you need files to persist between sessions, you'll need to integrate cloud storage like:
- AWS S3
- Google Cloud Storage
- Streamlit's experimental connection to databases

This requires additional setup and is beyond the scope of this basic deployment.

## Usage

1. Navigate to the "Reference Check" page
2. Upload your CPM document
3. Enter the SFC number
4. Click "Save CPM & Extract IDs"
5. Download the extracted IDs as Excel or copy the text list
6. Optionally save to the CPM Library for quick access (temporary)

## Troubleshooting

**App won't start:**
- Check the logs in Streamlit Community Cloud
- Verify all files are uploaded correctly
- Ensure `requirements.txt` has all dependencies

**Extraction not working:**
- Enable "Debug Mode" to see detailed processing
- Check that your CPM format is supported (PDF, DOCX, DOCM)
- Verify checkboxes are properly formatted in the document

**Files disappearing from library:**
- This is expected behavior on cloud deployment
- Always download extracted IDs immediately
- Consider implementing persistent storage for production use

## Support

For issues or questions, contact your IT administrator or the application developer.

## Version

Current version: 4.0 (Cloud)

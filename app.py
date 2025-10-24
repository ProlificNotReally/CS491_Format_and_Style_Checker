# app.py - Quick Flask app for autocorrecting text or docs. Threw this together for a demo project for our capstone.
# By: Omar Hassan
# Last tweak: Oct 9, 2025 - added some extra checks for empty files.

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from spellchecker import SpellChecker  # Grabbed this lib for basic spellchecking
from werkzeug.utils import secure_filename
from io import BytesIO
import re
import logging  # Added logging to help us out in dev when needed

# Doc parsing libraries - need these for PDFs and Word files
from PyPDF2 import PdfReader  # pip install PyPDF2 if you haven't
from docx import Document     # pip install python-docx both of these needed before these commands in here can work.

app = Flask(__name__)
app.secret_key = 'secret-dev-key'  # Change this for real prod or when we're submitting the project.

# Set up some basic logging - this just logs to the console for now important for later usage
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Spellchecker setup - English only since it's what we need. 
spell = SpellChecker(language='en')

# Regex for grabbing words - handles special grammar like apostrophes like "don't"
word_pattern = re.compile(r"[A-Z a-z']+")

def fix_spelling(text):
    # Helper func to swap misspelled words while keeping caps
    def swap_match(match):
        word = match.group(0)
        lower_word = word.lower()
        if lower_word in spell:  # If it's fine, leave it alone
            return word
        fix = spell.correction(lower_word) or lower_word  # Fallback to word if no suggestion comes up.
        # Keep the same cases: all caps, title case, or lowercase
        if word.isupper():
            return fix.upper()
        elif word[0].isupper():
            return fix.capitalize()
        return fix
    
    return word_pattern.sub(swap_match, text)

# Pull all of the text from PDF - loop through pages and grab whatever is there.
def get_text_from_pdf(file_obj):
    try:
        file_obj.stream.seek(0)  # Reset stream just in case something goes wrong.
        reader = PdfReader(file_obj.stream)
        text_bits = []
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:  # Skip empty pages
                text_bits.append(page_text)
        return '\n'.join(text_bits)
    except Exception as e:
        logger.error(f"Oops, PDF read failed: {e}")
        return ''

# Get text from DOCX - simpler, just paragraphs
def get_text_from_docx(file_obj):
    try:
        file_obj.stream.seek(0)
        doc = Document(file_obj.stream)
        return '\n'.join(p.text for p in doc.paragraphs if p.text.strip())  # Ignore empty paragraphs 
    except Exception as e:
        logger.error(f"Docx extraction bombed: {e}")
        return ''

# Main page - just renders the template
@app.route('/', methods=['GET'])
def home():
    return render_template('index.html')

# Handle text input from form
@app.route('/autocorrect-text', methods=['POST'])
def handle_text_correction():
    input_text = request.form.get('text', '').strip()
    if not input_text:
        flash('Hey, need some text to work with!')
        return redirect(url_for('home'))
    corrected_text = fix_spelling(input_text)
    # Pass back original and corrected for the template to show
    return render_template('index.html', original=input_text, corrected=corrected_text)

#  this portion here Handles file uploads
@app.route('/autocorrect-file', methods=['POST'])
def handle_file_correction():
    uploaded_file = request.files.get('document')
    if not uploaded_file or not uploaded_file.filename:
        flash('No file selected, try again.')
        return redirect(url_for('home'))
    
    safe_name = secure_filename(uploaded_file.filename)
    file_ext = safe_name.lower().split('.')[-1] if '.' in safe_name else ''
    
    if file_ext == 'pdf':
        raw_text = get_text_from_pdf(uploaded_file)
    elif file_ext == 'docx':
        raw_text = get_text_from_docx(uploaded_file)
    elif file_ext == 'txt':
        uploaded_file.stream.seek(0)
        raw_text = uploaded_file.stream.read().decode('utf-8', errors='ignore')
    else:
        flash('Only .pdf, .docx, or .txt files, please.')
        return redirect(url_for('home'))
    
    if not raw_text.strip():
        flash('File looks empty or no text found.')
        return redirect(url_for('home'))
    
    corrected_text = fix_spelling(raw_text)
    
    # This will send back as a downloadable TXT
    output_bytes = BytesIO(corrected_text.encode('utf-8'))
    output_filename = f'fixed_{safe_name.rsplit(".", 1)[0]}.txt'
    return send_file(output_bytes, download_name=output_filename, as_attachment=True, mimetype='text/plain')

if __name__ == "__main__":
    app.run(debug=True)
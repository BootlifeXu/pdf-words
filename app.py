import os, secrets, tempfile
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from processor import convert_pdf_to_docx
app=Flask(__name__); app.secret_key=os.environ.get('SECRET_KEY',secrets.token_hex(32)); MAX_UPLOAD_MB=int(os.environ.get('MAX_UPLOAD_MB','50')); app.config['MAX_CONTENT_LENGTH']=MAX_UPLOAD_MB*1024*1024

def allowed(f): return '.' in f and f.rsplit('.',1)[1].lower()=='pdf'
@app.get('/')
def index(): return render_template('index.html',max_upload_mb=MAX_UPLOAD_MB)
@app.errorhandler(413)
def too_large(e): flash(f'File is too large. Maximum size is {MAX_UPLOAD_MB} MB.','error'); return redirect(url_for('index'))
@app.post('/convert')
def convert():
    uploaded=request.files.get('pdf')
    if not uploaded or not uploaded.filename: flash('Please choose a PDF file.','error'); return redirect(url_for('index'))
    if not allowed(uploaded.filename): flash('Only PDF files are accepted.','error'); return redirect(url_for('index'))
    workdir=tempfile.mkdtemp(prefix='pdfword_'); input_path=os.path.join(workdir,secure_filename(uploaded.filename))
    try:
        uploaded.save(input_path)
        result=convert_pdf_to_docx(input_path,workdir,request.form.get('page_breaks')=='on',request.form.get('smart_page_breaks')=='on')
        return send_file(result['docx_path'],as_attachment=True,download_name=result['docx_name'],mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as exc:
        app.logger.exception('Conversion failed'); flash(f'Conversion failed: {exc}','error'); return redirect(url_for('index'))
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT','5000')),debug=False)

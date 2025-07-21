import os
import time
import pythoncom
import win32com.client
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def merge_ppts_using_com(files):
    # Required for COM use in Flask/threaded context
    pythoncom.CoInitialize()

    # Start PowerPoint COM instance
    ppt_instance = win32com.client.Dispatch("PowerPoint.Application")
    ppt_instance.Visible = 1

    # Open the first presentation
    base_ppt = ppt_instance.Presentations.Open(os.path.abspath(files[0]), WithWindow=False)

    # Insert slides from the rest
    for i in range(1, len(files)):
        source_path = os.path.abspath(files[i])
        if os.path.exists(source_path):
            base_ppt.Slides.InsertFromFile(source_path, base_ppt.Slides.Count)

    # Save the merged presentation with a unique name
    output_path = os.path.abspath(f"merged_{int(time.time())}.pptx")
    base_ppt.SaveAs(output_path)

    # Clean up
    base_ppt.Close()
    ppt_instance.Quit()
    pythoncom.CoUninitialize()

    return output_path


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('ppt_files')
        if not files:
            return "No files uploaded", 400

        ppt_paths = []
        for file in files:
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            ppt_paths.append(path)

        try:
            output_path = merge_ppts_using_com(ppt_paths)
        except Exception as e:
            return f"An error occurred: {e}", 500
        finally:
            # Clean up uploaded files
            for path in ppt_paths:
                if os.path.exists(path):
                    os.remove(path)

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)

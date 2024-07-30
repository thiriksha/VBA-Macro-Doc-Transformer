import google.generativeai as genai
import time
from flask import Flask, request, render_template
import zipfile
import os
import shutil
from oletools.olevba import VBA_Parser
import sys
from graphviz import Digraph

# Configure Google Generative AI
genai.configure(api_key=os.getenv("GENAI_API_KEY"))  # Use environment variable for API key

# Set up the model
generation_config = {
    "temperature": 0.9,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 2048,
}
model = genai.GenerativeModel(model_name="gemini-1.0-pro", generation_config=generation_config)

app = Flask(__name__)  # Corrected from _name_ to __name__

def extract_vba_code(file_path):
    vba_code = ""
    print("Starting to extract VBA code...")
    sys.stdout.flush()
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall('temp')
        print("Files extracted to 'temp' directory")
        sys.stdout.flush()
    except Exception as e:
        print(f"Error extracting files: {e}")
        sys.stdout.flush()
        return "", ""

    for root, dirs, files in os.walk('temp'):
        for file in files:
            print(f"Found file: {file}")
            sys.stdout.flush()
            if file == 'vbaProject.bin':
                print("Found vbaProject.bin, parsing...")
                sys.stdout.flush()
                try:
                    vba_parser = VBA_Parser(os.path.join(root, file))
                    if vba_parser.detect_vba_macros():
                        for (filename, stream_path, vba_filename, vba_code_content) in vba_parser.extract_macros():
                            print(f"Extracting macro: {vba_filename}")
                            sys.stdout.flush()
                            vba_code += f"Filename: {vba_filename}\n"
                            vba_code += vba_code_content + "\n\n"
                    else:
                        print("No VBA macros detected.")
                        sys.stdout.flush()
                except Exception as e:
                    print(f"Error parsing vbaProject.bin: {e}")
                    sys.stdout.flush()

    try:
        shutil.rmtree('temp')
        print("Temporary files deleted.")
        sys.stdout.flush()
    except Exception as e:
        print(f"Error deleting temporary files: {e}")
        sys.stdout.flush()

    print("Extraction complete.")
    sys.stdout.flush()
    return vba_code

def get_code_explanation(vba_code):
    try:
        convo = model.start_chat(history=[])
        convo.send_message(f"Explain the following VBA code in 2 or 3 small paragraphs:\n\n{vba_code}")
        time.sleep(2)  # Small delay to ensure response
        result = convo.last.text
        return result
    except Exception as e:
        print(f"Error getting code explanation: {e}")
        return "Error generating explanation."

def get_code_explanation_flow_chart(vba_code):
    try:
        convo = model.start_chat(history=[])
        convo.send_message(f"List out the main functionality of the code in Module1.bas in short, numbered steps suitable for creating a flowchart. Focus on the functionality and avoid technical details. No other extra statements. Just list out the steps only:\n\n{vba_code}")
        time.sleep(2)  # Small delay to ensure response
        result = convo.last.text
        return result
    except Exception as e:
        print(f"Error getting code explanation for flowchart: {e}")
        return "Error generating flowchart steps."

def create_process_flow_diagram(explanation):
    steps = explanation.split('\n')
    steps = [step.strip() for step in steps if step.strip()]

    dot = Digraph(comment='VBA Macro Process Flow')
    dot.attr(rankdir='TB')  # Change direction to top-to-bottom

    for i, step in enumerate(steps):
        dot.node(f'Step {i+1}', step)

    file_path = 'static/vba_macro_process_flow'
    dot.render(file_path, format='png', cleanup=True)
    return 'vba_macro_process_flow.png'

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/analyze', methods=['POST'])
def analyze_file():
    file = request.files['file']
    file_path = "uploaded_file.xlsm"
    try:
        file.save(file_path)
        vba_code = extract_vba_code(file_path)
        
        if vba_code:
            explanation = get_code_explanation(vba_code)
            explanation_flow = get_code_explanation_flow_chart(vba_code)
            diagram_path = create_process_flow_diagram(explanation_flow)
        else:
            explanation = "No VBA code extracted."
            diagram_path = None
        
        os.remove(file_path)
    except Exception as e:
        print(f"Error processing file: {e}")
        vba_code = "An error occurred during file processing."
        explanation = "Unable to generate explanation."
        diagram_path = None

    return render_template('result.html', vba_code=vba_code, explanation=explanation, diagram_path=diagram_path)

if __name__ == '__main__':
    if not os.path.exists('static'):
        os.makedirs('static')
    app.run(debug=True)

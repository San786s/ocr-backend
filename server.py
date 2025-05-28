# from flask import Flask, request, jsonify
# from flask_cors import CORS
# import pytesseract
# from PIL import Image
# import os

# app = Flask(__name__)
# CORS(app)  # Enable CORS for frontend-backend communication

# # Set the correct path for Windows
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# # Create an 'uploads' folder if it doesn't exist
# UPLOAD_FOLDER = "uploads"
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# @app.route("/extract_text", methods=["POST"])
# def extract_text():
#     if "image" not in request.files:
#         return jsonify({"error": "No image file uploaded"}), 400  # Return an error if no file is found

#     file = request.files["image"]
#     filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
#     file.save(filepath)  # Save the uploaded image

#     try:
#         img = Image.open(filepath)
#         extracted_text = pytesseract.image_to_string(img)  # Extract text using Tesseract
#         return jsonify({"text": extracted_text})  # Return extracted text in JSON format
#     except Exception as e:
#         return jsonify({"error": str(e)}), 500  # Return error if something goes wrong

# if __name__ == "__main__":
#     app.run(debug=True)

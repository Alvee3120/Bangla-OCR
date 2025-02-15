#
# import os
# import google.generativeai as genai
#
# # Configure API Key
# api = "Enter Your API key"  # Ensure your API key is valid
# genai.configure(api_key=api)
#
# # Function to upload a file
# def upload_file(path, mime_type=None):
#     file = genai.upload_file(path, mime_type=mime_type)
#     return file
#
# # Function to process all images in a folder
# def process_images(folder_path):
#     generation_config = {
#         "temperature": 1,
#         "top_p": 0.95,
#         "top_k": 40,
#         "max_output_tokens": 8192,
#         "response_mime_type": "text/plain",
#     }
#
#     model = genai.GenerativeModel(
#         model_name="gemini-2.0-flash",
#         generation_config=generation_config,
#     )
#
#     # Supported image formats
#     valid_extensions = {".jpg", ".jpeg", ".png", ".gif"}
#
#     # Iterate through all files in the folder
#     for file_name in os.listdir(folder_path):
#         file_path = os.path.join(folder_path, file_name)
#
#         # Check if the file is an image
#         if os.path.isfile(file_path) and os.path.splitext(file_name)[1].lower() in valid_extensions:
#             file = upload_file(file_path, mime_type="image/png")  # Change mime_type if needed
#
#             # Generate response
#             response = model.generate_content([
#                 file,
#                 "Extract the text from the image. no need to write any additional things"
#             ])
#
#             # Print result
#             print(f"{file_name} ---> {response.text}")
#
# # Provide the folder path here
# folder_path = r"C:\Users\User\Downloads\python test"  # Change this to your actual folder path
# process_images(folder_path)


import os
import pandas as pd
import google.generativeai as genai

# Configure API Key
api = "c"  # Ensure your API key is valid
genai.configure(api_key=api)

# Function to upload a file
def upload_file(path, mime_type=None):
    file = genai.upload_file(path, mime_type=mime_type)
    return file

# Function to process images and save results
def process_images(folder_path, output_file):
    generation_config = {
        "temperature": 1,
        "top_p": 0.95,
        "top_k": 40,
        "max_output_tokens": 8192,
        "response_mime_type": "text/plain",
    }

    model = genai.GenerativeModel(
        model_name="gemini-2.0-flash",
        generation_config=generation_config,
    )

    # Supported image formats
    valid_extensions = {".jpg", ".jpeg", ".png", ".gif"}

    # List to store results
    results = []

    # Iterate through all files in the folder
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        # Check if the file is an image
        if os.path.isfile(file_path) and os.path.splitext(file_name)[1].lower() in valid_extensions:
            file = upload_file(file_path, mime_type="image/png")  # Change mime_type if needed

            # Generate response
            response = model.generate_content([
                file,
                "Extract only the text from the image. No need to write any additional things"
            ])

            extracted_text = response.text.strip()  # Clean text

            # Debugging: Print output
            print(f"Processing: {file_name} ---> Extracted Text: {extracted_text}")

            # Append result to list
            results.append({"Sl no": len(results) + 1, "Image Name": file_name, "Image Text": extracted_text})

    # Save results to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"\nResults saved to {output_file}")

# Provide the folder path here
folder_path = r"C:\Users\User\Downloads\python test\new"  # Change this to your actual folder path
output_file = r"C:\Users\User\Downloads\pythontext.xlsx"  # Output Excel file

process_images(folder_path, output_file)

import os
import pyttsx3
from docx import Document

# Initialize text-to-speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)  # Adjust speech rate

# Folder containing .docx files
folder_path = 'C:/Users/Anish/Desktop/text to audio/docx_files'

# Output folder for audio files
output_folder = 'output_audio'
os.makedirs(output_folder, exist_ok=True)

# Iterate through all .docx files
for filename in os.listdir(folder_path):
    if filename.endswith('.docx'):
        docx_path = os.path.join(folder_path, filename)
        doc = Document(docx_path)
        
        # Extract full text
        full_text = '\n'.join([para.text for para in doc.paragraphs])

        # Set output file path
        audio_filename = os.path.splitext(filename)[0] + '.mp3'
        audio_path = os.path.join(output_folder, audio_filename)

        # Save audio
        engine.save_to_file(full_text, audio_path)
        print(f"Saved: {audio_path}")

# Finalize audio generation
engine.runAndWait()

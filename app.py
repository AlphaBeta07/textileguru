from gtts import gTTS
from docx import Document

# Load text from Word (.docx) file
def load_docx_text(file_path):
    doc = Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

# Path to your Word file
file_path = "06. Physical and Chemical Properties of Jute Fibre.docx"
text = load_docx_text(file_path)

# Convert text to audio
tts = gTTS(text, lang='en')
tts.save("06. Physical and Chemical Properties of Jute Fibre.mp3")

print("Audio file saved as '02. Cohesive Energy Density.mp3'")

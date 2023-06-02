import sys
import os
import pyttsx3
from docx import Document


import pyttsx3

def convert_docx_to_audio(docx_path,i, output_folder=None):
    # Load the DOCX file
    doc = Document(docx_path)

    # Extract text from paragraphs
    paragraphs = [p.text for p in doc.paragraphs]
    text = '\n'.join(paragraphs)

    # Initialize the text-to-speech engine
    engine = pyttsx3.init()

    # Get the available voices
    voices = engine.getProperty('voices')
    
    # Choose a voice (you can select a different index if desired)
    voice_index = i
    engine.setProperty('voice', voices[voice_index].id)

    # Set the output audio file path
    output_file = os.path.join(output_folder, os.path.basename(docx_path).replace('.docx', '.mp3')) if output_folder else \
        os.path.splitext(docx_path)[0] + '.mp3'

    # Convert text to audio and save the file
    engine.save_to_file(text, output_file)
    engine.runAndWait()


if __name__ == '__main__':
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Usage: python docx_to_audio.py <docx_file_path> [output_folder_path]")
        sys.exit(1)

    docx_file_path = sys.argv[1]
    output_folder_path = sys.argv[2] if len(sys.argv) == 3 else None

    if not os.path.isfile(docx_file_path):
        print("Invalid DOCX file path!")
        sys.exit(1)

    if output_folder_path and not os.path.isdir(output_folder_path):
        print("Invalid output folder path!")
        sys.exit(1)
    i = int(input('what voice to use 0-1: '))
    convert_docx_to_audio(docx_file_path, i, output_folder_path)
    print("Conversion complete!")

    
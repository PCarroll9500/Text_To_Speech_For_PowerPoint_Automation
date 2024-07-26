import openai
import argparse
import sys
import os

from API_Key import your_openai_key

# Initialize OpenAI client with API key
client = openai.OpenAI(api_key=your_openai_key)

def generate_speech(text, filename):
    try:
        response = client.audio.speech.create(
            model="tts-1",
            voice="echo",  # other voices: 'echo', 'fable', 'onyx', 'nova', 'shimmer'
            input=text
        )
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"MP3 file '{filename}' created successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    parser = argparse.ArgumentParser(description="Generate MP3 from text using OpenAI.")
    parser.add_argument('text', type=str, help='Text to convert to speech')
    parser.add_argument('filename', type=str, help='Output MP3 filename')
    args = parser.parse_args()
    
    print(f"Received text: {args.text}")
    print(f"Output filename: {args.filename}")
    
    generate_speech(args.text, args.filename)
    
    if os.path.exists(args.filename):
        print(f"File '{args.filename}' exists after creation.")
    else:
        print(f"File '{args.filename}' does not exist after creation.")

if __name__ == "__main__":
    main()

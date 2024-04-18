import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# API endpoint
api_url = 'https://api.openai.com/v1/completions'

# API Key
api_key = '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'

headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {api_key}'
}

# Read questions from Excel file
input_excel_path = "C:\\Users\\likit\\Downloads\\Travel_Destination_Answers (11).xlsx"
questions_df = pd.read_excel(input_excel_path)

def generate_response(question):
    data = {
        "model": "gpt-3.5-turbo-instruct",
        "prompt": f"{question} Provide a concise and informative response in up to 350 words with bullet points:\n•",
        "temperature": 0.7,
        "max_tokens": 1500  # Approximately 350 words
    }

    try:
        response = requests.post(api_url, headers=headers, json=data)
        if response.status_code == 200:
            result = response.json()
            # Format response into bullet points
            gpt_answer = result['choices'][0]['text'].strip()
            # Replace newlines with dot bullet points
            gpt_answer = gpt_answer.replace('\n', '\n')
            return gpt_answer
        else:
            print(f"Failed to get response for question: {question}")
            return None
    except Exception as e:
        print(f"An error occurred for question: {question}. Error: {e}")
        return None

# List to store generated responses
responses = []

# Use ThreadPoolExecutor for concurrent processing
with ThreadPoolExecutor(max_workers=5) as executor:
    # Process each question concurrently
    futures = [executor.submit(generate_response, question) for question in questions_df['Questions']]
    for future in futures:
        try:
            response = future.result()
            if response:
                responses.append(response)
        except Exception as e:
            print(f"An error occurred while processing a question: {e}")

# Create DataFrame from responses
df = pd.DataFrame({'Response': responses})

# Save DataFrame to Excel file
output_excel_path = "C:\\Users\\likit\\Downloads\\Travel_Destination_Answers (11).xlsx"
df.to_excel(output_excel_path, index=False)

print(f"Responses saved to {output_excel_path}")

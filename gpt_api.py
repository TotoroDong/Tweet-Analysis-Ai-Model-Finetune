import openai
import pandas as pd

# Api Key https://platform.openai.com/docs/overview
openai.api_key = 'your-openai-api-key'

def analyze_tweet(tweet):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",  
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Analyze this tweet: {tweet}"}
            ]
        )
        return response['choices'][0]['message']['content'].strip()
    except Exception as e:
        return f"Error: {str(e)}"

def process_tweets(input_excel, output_excel, tweet_column, result_column):
    # Read the Excel file
    df = pd.read_excel(input_excel)

    # Analyze each tweet and store the result in a new column
    df[result_column] = df[tweet_column].apply(analyze_tweet)

    # Save the results to a new Excel file
    df.to_excel(output_excel, index=False)
    print(f"Analysis complete. Results saved to {output_excel}")

# Example usage
input_excel = 'tweets.xlsx'  #  input Excel file path
output_excel = 'analyzed_tweets.xlsx'  #  desired output Excel file path
tweet_column = 'Tweet'  # Replace with the column name containing tweets
result_column = 'Analysis'  # Name of the new column where results will be stored

process_tweets(input_excel, output_excel, tweet_column, result_column)

import pandas as pd
import requests
from bs4 import BeautifulSoup

def read_universities_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df
def get_qs_ranking(university_name):
    search_url = f"https://www.topuniversities.com/universities/{university_name.replace(' ', '-')}"    

    try:
        response = requests.get(search_url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            ranking = soup.find('div', class_='ranking-rank').get_text(strip=True)
            return ranking
        else:
            return "Not Found"
    except Exception as e:
        return f"Error: {str(e)}"

def add_qs_rankings(df):
    df['QS Ranking'] = df['University'].apply(get_qs_ranking)
    return df


def write_universities_to_excel(df, output_file_path):
    df.to_excel(output_file_path, index=False)

# Main script
input_file = 'universities.xlsx'
output_file = 'universities_with_qs_rankings.xlsx'

# Lettura
universities_df = read_universities_from_excel(input_file)

universities_with_rankings_df = add_qs_rankings(universities_df)

write_universities_to_excel(universities_with_rankings_df, output_file)

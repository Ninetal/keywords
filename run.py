import xlsxwriter
from string import punctuation, digits
from pymystem3 import Mystem
import nltk
from nltk.corpus import stopwords
nltk.download('stopwords')

KEYWORDS_COUNT = 500
OUTPUT_FILE_NAME = 'output.xlsx'


def get_file() -> str:
    file_path = input("\nEnter file path: ")
    if not file_path:
        raise Exception('File path is empty')
    with open(file_path, 'r', encoding='utf-8') as file:
        data = file.read()
    return data


def identify_top_words(all_words: list) -> list:
    freq_dist = nltk.FreqDist(all_words)
    return freq_dist.most_common(KEYWORDS_COUNT)


def tokenize(data: str) -> list:
    mystem = Mystem()
    garbage = ['что', 'который', 'это', 'вот', '—', '–', '...']
    stop_words = stopwords.words('russian')

    tokens = mystem.lemmatize(data.lower())

    tokens = [i.replace(" ", "") for i in tokens]
    tokens = [i for i in tokens if (i not in punctuation
                                    and i not in digits
                                    and i not in garbage
                                    and len(i) > 2)]

    tokens = [i for i in tokens if (i not in stop_words)]
    return tokens


def xls_writer(content: list):
    workbook = xlsxwriter.Workbook(OUTPUT_FILE_NAME)
    worksheet = workbook.add_worksheet()
    row = 0
    column = 0
    for item in content:
        worksheet.write(row, column, item)
        row += 1
    workbook.close()
    print('Result in {}'.format(OUTPUT_FILE_NAME))


def find_keywords():
    data = get_file()
    tokens = tokenize(data)
    top = identify_top_words(tokens)
    top = [item[0] for item in top]
    xls_writer(top)


if __name__ == "__main__":
    find_keywords()

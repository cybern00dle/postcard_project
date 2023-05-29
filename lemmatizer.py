import re
from string import punctuation

import pandas as pd
from pymystem3 import Mystem


class Lemmatizer:
    """
    Cleans and lemmatizes postcard texts
    """
    def __init__(self, text: str, punctuation: str):
        self.raw_text = str(text)
        self.punctuation = punctuation
        self.clean_text = ''
        self.analyzer = Mystem()
        self.lemmas = []

    def clean(self) -> None:
        """
        Cleans the given text
        :return: None
        """
        clean_text = ''
        for i in self.raw_text.lower().replace('\n', ' ').strip():
            if i not in self.punctuation:
                clean_text += i
        self.clean_text = clean_text

    def lemmatize(self) -> None:
        """
        Lemmatizes clean tokens
        :return: None
        """
        self.lemmas = self.analyzer.lemmatize(self.clean_text)

    def run(self) -> str:
        """
        Runs the process of cleaning, tokenization and lemmatization
        :return: a string with lemmatized text
        """
        self.clean()
        self.lemmatize()
        return ' '.join(self.lemmas)


def main() -> None:
    corpus = pd.read_excel('corpus.xlsx')
    needed_postcards = pd.DataFrame()
    for idx, text in enumerate(corpus['Текст открытки']):
        lemmatizer = Lemmatizer(text, punctuation)
        lemmas = lemmatizer.run()
        if re.search(r'скучать\s+за', lemmas):
            needed_postcards.loc[len(needed_postcards)] = corpus[idx]
    needed_postcards.to_excel('postcards.xlsx', index=False)


if __name__ == '__main__':
    main()

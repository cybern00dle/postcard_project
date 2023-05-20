import pandas as pd
from pymystem3 import Mystem
from string import punctuation


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
        # print необходим на этапе проверки правильности работы кода
        print(' '.join(self.lemmas))
        return ' '.join(self.lemmas)


def main() -> None:
    corpus = pd.read_excel('corpus.xlsx')
    index = 0
    for text in corpus['Текст открытки']:
        lemmatizer = Lemmatizer(text, punctuation)
        lemmas = lemmatizer.run()
        corpus.loc[index, 'Текст открытки'] = lemmas
        index += 1
    corpus.to_excel('corpus_lemmatized.xlsx', index=False)


if __name__ == '__main__':
    main()

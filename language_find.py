# -*- coding: utf-8 -*-
"""
Created on Mon Jul  3 16:04:48 2017
Purpose: Example for detecting language using a stopwords based approach
@author: AnkurDixit
"""

try:
    from nltk import wordpunct_tokenize
#    from nltk.corpus import stopwords
    from nltk.corpus import stop_words #Download and install this module via python command line
except ImportError:                      # pip install stop-words, then copy stop-words and paste it 
    print("The packages of NLTK is not installed")   #under the location as nltk.corpus



#----------------------------------------------------------------------
def _calc_lang_ratios(text):
    """
    Calculate probability of given text to be written in several languages and
    return a dictionary that looks like {'french': 2, 'spanish': 4, 'english': 0}
    
    @param text: Text whose language want to be detected
    @type text: str
    
    @return: Dictionary with languages and unique stopwords seen in analyzed text
    @rtype: dict
    """

    lang_ratio = {}

    '''
    nltk.wordpunct_tokenize() splits all punctuations into separate tokens
    
    >>> wordpunct_tokenize("That's thirty minutes away. I'll be there in ten.")
    ['That', "'", 's', 'thirty', 'minutes', 'away', '.', 'I', "'", 'll', 'be', 'there', 'in', 'ten', '.']
    '''

    tokens = wordpunct_tokenize(text)
    words = [word.lower() for word in tokens]

    # Compute per language included in nltk number of unique stopwords appearing in analyzed text
#    for language in stopwords.fileids():  #list(stop_words.LANGUAGE_MAPPING.keys())
    for language in list(stop_words.LANGUAGE_MAPPING.keys()):
#        stopwords_set = set(stopwords.words(language)) # set(stop_words.get_stop_words(language))
        stopwords_set = set(stop_words.get_stop_words(language))
        words_set = set(words)
        common_elements = words_set.intersection(stopwords_set)

        lang_ratio[language] = len(common_elements) # language "score"
        
    print("\n\nThe Language ratios are :", lang_ratio)    
    return lang_ratio


#----------------------------------------------------------------------
def detect_language(text):
    """
    Calculate probability of given text to be written in several languages and
    return the highest scored.
    
    It uses a stopwords based approach, counting how many unique stopwords
    are seen in analyzed text.
    
    @param text: Text whose language want to be detected
    @type text: str
    
    @return: Most scored language guessed
    @rtype: str
    """

    ratios = _calc_lang_ratios(text)

    most_rated_language = max(ratios, key=ratios.get)

    return most_rated_language



if __name__=='__main__':
    
    term = "Please enter the text for which Language recognisation is required : " #term we want to search for
    print(term)
    text = input() #read input from user
    language = detect_language(text)

    print("\n\nThe Language is :",language)
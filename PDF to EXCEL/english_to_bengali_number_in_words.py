from num2words import num2words
from googletrans import Translator

def english_to_bengali_number_in_words(english_number):
    # Convert English number to words using Indian numbering system
    words_in_english = num2words(english_number, lang='en_IN')

    # Translate to Bengali
    translator = Translator()
    words_in_bengali = translator.translate(words_in_english, dest='bn').text

    # Remove commas and add "টাকা মাত্র" at the end
    modified_output = words_in_bengali.replace(',', '') + " টাকা মাত্র"

    return modified_output

# Example usage:
string = (str(0143.34) + "").split('.')[0]
english_number = float(string)
english_number = english_number
bengali_words = english_to_bengali_number_in_words(english_number)
print(bengali_words)


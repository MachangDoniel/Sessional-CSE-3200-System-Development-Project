from googletrans import Translator

def convert_year_term_suffixes_to_bengali(text):
    # Dictionary mapping English year_term_suffixes to Bengali
    year_term_suffixes_mapping = {
        "1st": "১ম",
        "2nd": "২য়",
        "3rd": "৩য়",
        "4th": "৪র্থ",  # You can add more mappings as needed
        # Add more mappings for other year_term_suffixes
    }

    # Replace English year_term_suffixes with Bengali equivalents
    for suffix in year_term_suffixes_mapping:
        if suffix in text:
            text = text.replace(suffix, year_term_suffixes_mapping[suffix])

    return text

# Example usage:
english_text = "1st"
bengali_text = convert_year_term_suffixes_to_bengali(english_text)
print(bengali_text)

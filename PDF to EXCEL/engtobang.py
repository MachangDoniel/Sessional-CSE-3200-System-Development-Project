from googletrans import Translator

def convert_suffix_to_bengali(text):
    # Dictionary mapping English suffixes to Bengali
    suffixes_mapping = {
        "1st": "১ম",
        "2nd": "২য়",
        "3rd": "৩য়",
        "4th": "৪র্থ",  # You can add more mappings as needed
        # Add more mappings for other suffixes
    }

    # Replace English suffixes with Bengali equivalents
    for suffix in suffixes_mapping:
        if suffix in text:
            text = text.replace(suffix, suffixes_mapping[suffix])

    return text

# Example usage:
english_text = "1st"
bengali_text = convert_suffix_to_bengali(english_text)
print(bengali_text)

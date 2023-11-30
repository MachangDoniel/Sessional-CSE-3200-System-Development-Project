from googletrans import Translator
import re

def should_skip_translation(text):
    name_patterns = [r'Dean', r'Md\.', r'Dr\.', r'Sk\.', r'Most\.', r'Fatema']
    for pattern in name_patterns:
        if re.search(pattern, text):
            return True
    return False

def translate_to_bengali(text):
    translator = Translator()

    # Translation rules for specific patterns
    translation_rules = {
        r'Dean': 'ডিন',
        r'Md\.': 'মোঃ',
        r'Dr\.': 'ড.',
        r'Sk\.': 'শেখ',
        r'Most': 'মোসাম্মৎ',
        r'Fatema': 'ফাতেমা'
    }

    # Translate only the parts that don't match predefined patterns or apply specific rules
    parts = text.split()
    translated_parts = []
    for part in parts:
        if not should_skip_translation(part):
            # Apply the specific translation rule if found
            for pattern, replacement in translation_rules.items():
                if re.search(pattern, part):
                    part = re.sub(pattern, replacement, part)
                    break
            translated_part = translator.translate(part, dest='bn').text
        else:
            # Use provided translation rules when skipping translation
            for pattern, replacement in translation_rules.items():
                if re.search(pattern, part):
                    translated_part = re.sub(pattern, replacement, part)
                    break
        translated_parts.append(translated_part)

    return ' '.join(translated_parts)

# Example usage:
english_text = "Dean"
bengali_text = translate_to_bengali(english_text)
print(f"English: {english_text}")
print(f"Bengali: {bengali_text}")

# ড. শেখ মোঃ মাসুদুল আহসান
# মোসাম্মৎ. ক্যানিজ ফাতেমা ইসা
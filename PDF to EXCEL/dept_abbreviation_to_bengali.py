from googletrans import Translator

class DepartmentTranslator:
    def __init__(self):
        self.translator = Translator()
        self.dept_suffixes_mapping = {
             
        "computer science and engineering": "সিএসই",
        "computer science & engineering": "সিএসই",
        "electrical and electronic engineering": "ইইই",
        "electrical & electronic engineering": "ইইই",
        "electronics and communication engineering": "ইসিই",
        "electronics & communication engineering": "ইসিই",
        "biomedical engineering": "বিএমই",
        "materials science and engineering": "এমএসই",
        "materials science & engineering": "এমএসই",
        "civil engineering": "পুরকৌশল",
        "urban and regional planning": "ইউআরপি",
        "urban & regional planning": "ইউআরপি",
        "building engineering and construction management": "বিইসিএম",
        "building engineering & construction management": "বিইসিএম",
        "architecture": "স্থাপত্য",
        "mathematics": "গণিত",
        "math": "গণিত",
        "chemistry": "রসায়ন",
        "physics": "পদার্থ",
        "humanities": "মানবিক",
        "mechanical engineering": "যন্ত্র প্রকৌশল",
        "industrial engineering and management": "শিল্প প্রকৌশল",
        "industrial engineering & management": "শিল্প প্রকৌশল",
        "energy science and engineering": "ইএসই",
        "energy science & engineering": "ইএসই",
        "leather engineering": "লেদার",
        "textile engineering": "টেক্সটাইল",
        "chemical engineering": "টেক্সটাইল",
        "mechatronics engineering": "মেকাট্রনিক্স",
        }

    def dept_translate_to_bengali(self, english_text):
        bengali_text = self.dept_suffixes_mapping.get(english_text.lower())
        if not bengali_text:
            # If the translation is not found in the mapping, use Google Translate
            translated = self.translator.translate(english_text, dest='bn')
            bengali_text = translated.text
        return bengali_text

# Example usage:
translator = DepartmentTranslator()
english_text = "Computer SciencE ANd Engineering"
bengali_text = translator.dept_translate_to_bengali(english_text.lower())
print(f"English: {english_text}")
print(f"Bengali: {bengali_text}")
import re
import os
import win32com.client as win32

WORD_REPLACEMENTS = {
    r'\b(organize)\b': 'organise',
    r'\b(eg)\b': 'for example'
}

mentioned_names = {}
last_name_usage = {}

def replace_words(text):
    """Replace specific terms based on the WORD_REPLACEMENTS dictionary, preserving case."""
    for pattern, replacement in WORD_REPLACEMENTS.items():
        matches = re.finditer(pattern, text, flags=re.IGNORECASE)
        for match in reversed(list(matches)):
            original = match.group(1)
            start, end = match.span(1)

            if original.islower():
                new_word = replacement.lower()
            elif original.istitle():
                new_word = replacement.capitalize()
            elif original.isupper():
                new_word = replacement.upper()
            else:
                new_word = replacement

            text = text[:start] + new_word + text[end:]
    return text

def format_names(text):
    """Format names with initials, and simplify repeated mentions intelligently."""
    text = re.sub(r'\b([A-Z])\s+([A-Z][a-z]+)\b', r'\1. \2', text)

    name_pattern = r'\b(Dr|Mr|Mrs|Ms|Prof|Sir|Lady|Capt|Lt|Gen|Col|Rev)\s+([A-Z][a-z]+)(?:\s+([A-Z])\s+)?(\s+[A-Z][a-z]+)\b'

    def simplify_name(match):
        title = match.group(1)
        first = match.group(2)
        middle = match.group(3) or ""
        last = match.group(4).strip()

        middle = f"{middle}. " if middle else ""
        full_name = f"{title} {first} {middle}{last}"
        short_name = f"{title} {last}"

        last_name_usage.setdefault(last, set()).add(first)

        if full_name in mentioned_names:
            return full_name if len(last_name_usage[last]) > 1 else short_name
        else:
            mentioned_names[full_name] = True
            return full_name

    return re.sub(name_pattern, simplify_name, text)

def process_text(text):
    """Process text outside of quotes: apply replacements and name formatting."""
    pattern = r'(".*?"|“.*?”)' # for smart quotes used by word
    segments = re.split(pattern, text) 
    processed = []

    for i, segment in enumerate(segments):
        if i % 2 == 0:  # Outside quotes
            segment = replace_words(segment)
            segment = format_names(segment)
        processed.append(segment)

    return ''.join(processed)

def correct_document(input_file, output_file):
    """Open a Word document, apply text corrections, and save a new version."""
    global mentioned_names, last_name_usage
    mentioned_names = {}
    last_name_usage = {}

    input_path = os.path.abspath(input_file)
    output_path = os.path.abspath(output_file)

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(input_path)

        for paragraph in doc.Paragraphs:
            original_text = paragraph.Range.Text
            corrected_text = process_text(original_text)
            paragraph.Range.Text = corrected_text

        doc.SaveAs(output_path)
        doc.Close()
    finally:
        word.Quit()

if __name__ == "__main__":
    correct_document("input.docx", "corrected.docx")

import textwrap
import pyperclip
import re

FILENAME = "./temp.md"
column_width = 72

# Open file
with open(FILENAME, 'r', encoding='utf-8') as file:
    file_text = file.read()

split_lines = file_text.split("\n\n")

markdown = []
for line in split_lines:
    # Create header and body
    # print(line)
    match = re.search(r"^(\d+\.\d*)(\s.+\n?)$", line)
    if match:
        section_number, body_text = match.groups()
        text = f"- **{section_number}** {body_text.strip()}"
    else:
        text = f"  - {line.strip()}"

    wrapped_text = textwrap.wrap(text,
                                 column_width,
                                 break_long_words=False,
                                 )
    body_text = "\n  ".join(wrapped_text)

    markdown.append(body_text)

# Output text
print('\n'.join(markdown))

pyperclip.copy('\n'.join(markdown))

print("Text copied to clipboard!")

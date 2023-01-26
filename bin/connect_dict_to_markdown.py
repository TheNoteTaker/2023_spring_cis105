import textwrap
import pyperclip

FILENAME = "../modules/module_02/temp.txt"
column_width = 72

# Open file
with open(FILENAME, 'r', encoding='cp1252') as file:
    file_text = file.read()

split_lines = file_text.split("\n\n")

markdown = []
for line in split_lines:
    # Create header and body
    header, body = line.split("\n", maxsplit=1)

    # Create word to be defined in Markdown format
    header_text = f"- **{header.strip()}:**"

    # Create final text
    merged_text = " ".join([header_text, body.strip().replace("\n", " ")])
    wrapped_text = textwrap.wrap(merged_text,
                                 column_width,
                                 break_long_words=False,
                                 )
    body_text = "\n  ".join(wrapped_text)

    markdown.append(body_text)

# Output text
print('\n'.join(markdown))

pyperclip.copy('\n'.join(markdown))

print("Text copied to clipboard!")

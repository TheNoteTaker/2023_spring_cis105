import subprocess

FILENAME = "./temp_definitions.md"

with open(FILENAME, 'r', encoding='utf-8') as file:
    definitions = "\n- ".join(sorted([line.lstrip() for
                                        line in file.read().split("\n-")]))

print(definitions)

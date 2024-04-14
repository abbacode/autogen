"""This file shows an example of how to load variables.xlsx into a dictionary."""
from autogen import (
    create_data_frame_from_excel_file,
    extract_variables_from_data_frame,
)

df = create_data_frame_from_excel_file(file_path="variables.xlsx")
variables = extract_variables_from_data_frame(data_frame=df)

print(variables["examples"])
print(variables["contact_info"][0]["name"])

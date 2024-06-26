# Overview

Autogen allows the creation of word documents and visio diagrams from templates. The templates rely on the Jinja2
template engine to render and create new documents based on variables stored in an external excel spreadsheet.

<br><img width="485" alt="word_template_example" src="https://github.com/abbacode/autogen/assets/13191198/5463f52b-de4a-4d13-8dc1-4f6d8c48d80b"><br>
<br><img width="485" alt="word_template_example" src="https://github.com/abbacode/autogen/assets/13191198/73cb5533-6975-40ea-be39-f2ac4da01e28">





## Why Use Autogen?

Consultants typically use these three Microsoft products when generating statement of works, design documents, as-built documents,
change plans, etc and it is a slow and tedious process that is often error prone.

For large and complex engagements were dozens or hundreds of sites are involved the problem is amplified.
Given that information is now distributed across multiple sources it is also difficult to automate as the data
quickly becomes out of sync.

With Autogen, consultants and engineers can leverage a single data source (Excel) for their engagements and use this
for all aspects of the project life cycle. From design to build to handover, all your data is now located in a single
source of truth and can be used to generate artefacts on demand using templates.

## External Requirements

- Microsoft Visio must be installed on the local machine.

## Quick Start

1. Create and activate the virtual environment
2. Install the requirements: `pip install -r requirements.txt`
3. Update the `template_design.docx` to suit your requirements
4. Update the `template_design.vsdx` to suit your requirements
5. Update the `variables.xlsx` to suit your requirements
6. Run the script: `python main.py`

The following artefacts will automatically be generated:

- detailed_design.docx
- detailed_design.vsdx

Example of the template:<br><br>
<img width="485" alt="word_template_example" src="https://github.com/abbacode/autogen/assets/13191198/195a16bb-bbd7-42cf-99ac-d08fd0fd5d92">
<br>
Example after it has been rendered:<br><br>
<img width="490" alt="word_rendered_example" src="https://github.com/abbacode/autogen/assets/13191198/925971e5-df23-4c0a-90b5-3f51bbf7cac0">


## Variables

The `variables.xlsx` file is where all the variable information is stored. You can structure and organise
your file using as many worksheets and tables as required. Each worksheet should contain a single table for ease 
of use, however, there is also support for multiple tables within the same worksheet. Examples for this will be shown
under the advanced usage section. 

These tables can then be referenced from inside the word document and visio diagram templates using jinja2 syntax.
For a primer on jinja2 usage refer to: https://ultraconfig.com.au/blog/jinja2-a-crash-course-for-beginners/

### Vertical Tables

Consider the following vertical table is stored in the `examples` worksheet:<br>

![image](https://github.com/abbacode/autogen/assets/13191198/e3474388-11ee-493e-92eb-2c267394aee2)
<br><br>
These types of tables are great at storing key/value data that can be easily referenced inside your word document template:<br>
|Syntax       | Value      |
|:----------- |----------:|
| {{ examples.key1 }}  | value1
| {{ examples.key2 }}  | value2

### Horizontal Tables
Consider the following horitzontal table is stored in the `contact_info` worksheet:<br>

![image](https://github.com/abbacode/autogen/assets/13191198/4d683d1c-6a97-4c1c-8639-9675c4f1bd44)
<br><br>
These types of tables store the data as a list of dictionaries and can be referenced inside your word document template:<br>
|Syntax       | Value      |
|:----------- |----------:|
| {{ contact_info[0].name }}  | Abdul
| {{ contact_info[1].address }}  | China
<br>
It's also possible to show the entire table inside a word document using jinja2 syntax: <br>
![image](https://github.com/abbacode/autogen/assets/13191198/d08a09bb-8423-4720-aec0-136d9ac53634)

## Advanced Usage

### How to have add multiple tables in a single worksheet
Multiple tables within the same worksheet are supported with each table being assigned
an automatically generated name using the following convention: 
<br>
> <worksheet_table_name>table_<table_number>

i.e. if the `more_examples` worksheet was created as follows:<br><br>
![image](https://github.com/abbacode/autogen/assets/13191198/4deb14df-a976-4af5-9f65-175e1310c069)
<br><br>
The key info for using multiple tables within the same worksheet:
- Mixed table types (vertical/horiztonal) can be used within the same worksheet
- For a table to be considered valid you must have at least two rows defined onto two consecutive lines.
- The last table on the page can ALSO be referenced using just the worksheet name
- Any table headers starting with a # will be ignored and can be used to add comments

Examples on how to reference this data inside your template:
<br><br>

|Syntax       | Value      |
|:----------- |----------:|
| {{ more_examples_table_1.key1 }}  | value1
| {{ more_examples_table_1.key2 }}  | value2
| {{ more_examples_table_2[0].name }}  | abdul
<br>


### How to auto generate visio diagrams

Autogen also covers the automatic generation of visio diagrams. 

Take the following scenario where ```template_design.vsdx``` is setup as follows:<br><br>
![image](https://github.com/abbacode/autogen/assets/13191198/778a9df5-b2f6-4675-8eff-d68e824190c8)
<br>
And the ```diagram_variables``` worksheet within the variables file is setup as follows:<br><br>
![image](https://github.com/abbacode/autogen/assets/13191198/da1301b9-430a-4b4c-9183-aedeb412b150) 
<br>
Will automatically produce the ```detailed_design.vsdx``` with the following output: <br><br>
![image](https://github.com/abbacode/autogen/assets/13191198/98ab8f6a-f847-483b-b1df-9bc7f8d2ea02)


### How to include diagrams in word documents

Diagrams can automatically be inserted into a word document by referencing the visio tab name.
For example if the visio tab is called ```physical``` then the diagram can be automatically imported
inside a word document template using the following syntax:<br><br>
```{{ images.physical }}```

### How can I use my variables for other automation tasks

If you're interested in using th data within the `variables.xlsx` file via Python for other automation tasks:
```
from autogen import (
    create_data_frame_from_excel_file,
    extract_variables_from_data_frame,
)

# Read the data
df = create_data_frame_from_excel_file(file_path="variables.xlsx")
variables = extract_variables_from_data_frame(data_frame=df)

# Vertical table example
print(variables["worksheet_name"]["column_name")

# Horiztonal table example (row 1)
print(variables["worksheet_name"][0]["column_name"])

# Horiztonal table example (row 2)
print(variables["worksheet_name"][1]["column_name"])

```
NOTE: An example can be found at [here](https://github.com/abbacode/autogen/blob/main/example.py)


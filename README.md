# Overview

Autogen allows the creation of word documents and visio diagrams from templates. The templates rely on the Jinja2
template engine to render and create new documents based on variables stored in an external excel spreadsheet.

## Why Use Autogen?

Consultants typically use these three products when generating statement of works, design documents, as-built documents,
change plans, etc and it is a slow and tedious process that is often error prone.

For large and complex engagements were dozens or hundreds of sites are involved the problem is amplified.
Given that information is now distributed across multiple sources it is also difficult to automate as the data
quickly becomes out of sync.

With Autogen, consultants and engineers can leverage a single data source (Excel) for their engagements and use this
for all aspects of the project life cycle. From design to build to handover, all your data is now located in a single
source of truth and can be used to generate artefacts on demand using templates.

## External Requirements

- Microsoft Visio must be installed on the local machine.

## How To Use

1. Create and activate the virtual environment
2. Install the requirements: pip install -r requirements.txt 
3. Update the template_design.docx
4. Update the template_design.vsdx
5. Update the variables.xlsx
6. Run the script: python main.py

The artefacts generated:

- detailed_design.docx
- detailed_design.vsdx

## Advanced Usage

TODO: update this section. Include examples of how to interact with visio, etc.

## Example Output
The variables examples worksheet tab: <br>
<img width="214" alt="variable_example_worksheet" src="https://github.com/abbacode/autogen/assets/13191198/06889a61-8adb-4d39-96b1-8724596105f9">

The variables contact_info tab:<br>
<img width="289" alt="variable_contact_info_worksheet" src="https://github.com/abbacode/autogen/assets/13191198/7606089d-c5eb-413a-8934-87c02ecda663">

The word document template:<br>
<img width="485" alt="word_template_example" src="https://github.com/abbacode/autogen/assets/13191198/195a16bb-bbd7-42cf-99ac-d08fd0fd5d92">

After it has been rendered:<br>
<img width="490" alt="word_rendered_example" src="https://github.com/abbacode/autogen/assets/13191198/925971e5-df23-4c0a-90b5-3f51bbf7cac0">


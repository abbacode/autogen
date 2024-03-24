"""Example of using the automated design generator."""

import os
import sys

from adg import (
    add_images_to_template_variables,
    close_and_save_running_visio_application,
    close_and_save_running_word_application,
    create_data_frame_from_excel_file,
    export_visio_diagrams_to_png,
    extract_variables_from_data_frame,
    generate_design_diagrams,
    generate_design_document,
    remove_png_files_from_path,
)
from docxtpl import DocxTemplate

# The names of the files to either read or generate
design_template_document_filename = "template_design.docx"
design_template_diagram_filename = "template_design.vsdx"
design_variable_file_name = "variables.xlsx"
design_document_filename = "detailed_design.docx"
design_diagram_visio_filename = "detailed_design.vsdx"

basedir = os.path.dirname(os.path.realpath(__file__))

# The location where the templates are located
design_variables_path = os.path.join(basedir, design_variable_file_name)
template_document_path = os.path.join(basedir, design_template_document_filename)
template_diagram_path = os.path.join(basedir, design_template_diagram_filename)

# The location to store the newly generated files
design_document_path = os.path.join(basedir, design_document_filename)
design_diagram_visio_path = os.path.join(basedir, design_diagram_visio_filename)


if __name__ == "__main__":
    # Make sure visio and word applications are closed before running this script to avoid resource sharing issues
    # when we invoke win32com to open the applications in the background to perform conversion functions
    close_and_save_running_visio_application()
    close_and_save_running_word_application()

    # Read the data from the variables spreadsheet and store it in a pandas dataframe as it's easy to work with
    df = create_data_frame_from_excel_file(file_path=design_variables_path)
    if not df:
        sys.exit(0)

    # Convert the pandas data frame structure into a dictionary that contains all the variable values that can be
    # used while rendering the template document into the final document.
    template_variables = extract_variables_from_data_frame(data_frame=df)

    # Read the template document file into memory
    doc_template = DocxTemplate(template_document_path)

    # Extract the diagram label and variable values that are used when rendering visio diagrams
    design_diagram_label_values = template_variables["diagram_labels"]
    design_diagram_variable_values = template_variables["diagram_variables"]

    # Using the template visio diagram, create a new visio diagram while replacing placeholder text
    design_diagrams_generated = generate_design_diagrams(
        visio_template_path=template_diagram_path,
        save_to_path=design_diagram_visio_path,
        diagram_variable_values=design_diagram_variable_values,
        diagram_label_values=design_diagram_label_values,
    )
    if not design_diagrams_generated:
        sys.exit(0)

    # Export all tabs inside the visio file into separate *.png files
    visio_file_exported_as_png = export_visio_diagrams_to_png(
        visio_diagram_path=design_diagram_visio_path, save_to_path=basedir
    )
    if not visio_file_exported_as_png:
        sys.exit(0)

    # In order to reference images inside the word document, they need to be added to the variables dictionary
    # using special object types. Auotmatically scan all *.png files in the base dir and make them available.
    # NOTE: they can be referenced inside the design template via {{ images.<visio_tab_name> }}
    variables_containing_images = add_images_to_template_variables(
        doc_template=doc_template, variables=template_variables, image_path=basedir
    )
    if not variables_containing_images:
        print("WARNING: No *.png files were located or are available for this template.")

    # Using the design template, render the variables and create a new design document that has all the values
    # subsituted with the correct output
    design_document_generated = generate_design_document(
        doc_template=doc_template, save_to_path=design_document_path, variables=variables_containing_images
    )

    # Clean up all the temporary *.png files as they are no longer needed
    all_png_files_removed = remove_png_files_from_path(image_path=basedir)
    if not all_png_files_removed:
        print(f"WARNING: was unable to remove all *.png files from: {basedir}")

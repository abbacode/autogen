"""Automated design generator."""

import os
import pathlib
import shutil
import sys
from typing import Any, Dict, List, Optional

import pandas as pd
import win32com.client
import win32com.client as win32
from docxtpl import DocxTemplate, InlineImage
from jinja2.exceptions import UndefinedError
from vsdx import VisioFile

# The following worksheets are only allowed to use one table, any empty rows will be ignored
worksheets_with_single_table_allowed = [
    "diagram_labels",
    "diagram_variables",
]


def create_data_frame_from_excel_file(file_path: str) -> Optional[Dict[str, pd.DataFrame]]:
    """Load the excel workbook including all worksheets into a pandas dataframe.

    Args:
        file_path: the full path to the excel file

    Returns:
        xls: the spreadsheet with all worksheets within a single pandas dataframe
    """
    xls: None | Dict[str, pd.DataFrame] = None
    try:
        # load all worksheets into a single data frame
        xls = pd.read_excel(file_path, sheet_name=None)
    except PermissionError as e:
        print(f"Unable to read data from: {file_path} - due to: {e}")
    return xls


def create_data_frame_for_each_table(data_frame: pd.DataFrame) -> List[pd.DataFrame]:
    """Splits a single data frame into multiple data frames, one for each table detected.

    Each worksheet starts off with a single data frame, regardless of whether there is one or multiple tables within.
    This is difficult to work with as each table might contain different headings. This function will scan the data
    frame and break it up into multiple smaller data frames, one for each table containing the correct headings.

    NOTE: Tables are detected each time a new empty row is found.

    Args:
        data_frame: the pandas data frame containing the content of a excel spreadsheet

    Returns:
        table_data_frames: a list of data frames, one for each table detected within the worksheet
    """
    # Find all rows that are empty
    m = data_frame.isna().all(axis=1)
    # Automatically split the data frame into multiple smaller data frames, one for each table detected
    data_frames_split_by_empty_row = [group for key, group in data_frame[~m].groupby(m.cumsum())]
    # When the data frame is split into multiple smaller data frames, the correct data is mostly carried across, however
    # the columns/headers are inheritted from the first table and likely reference non existent columns. Pandas will
    # automatically try to read data from these columns resulting in the addition of junk data into the smaller data
    # frames. To prevent this from happening, we extract the headers manually and use them to re-create the data frames
    # properly while scrubing incorrect data.
    table_data_frames: List[pd.DataFrame] = []
    print("-------------------------------------------------------------------")
    table_no = 1
    for df in data_frames_split_by_empty_row:
        # Replace any empty cells in the table with an empty string, otherwise by default they would appear as NaN
        df.fillna("", inplace=True)
        # The first table does not require the same modification as subsequent tables in the worksheet.
        # We only need to drop Unnamed columns (these appear if you have multiple tables in a worksheet with different
        # column names) and columns starting with # as these are considered comment columns.
        if table_no == 1:
            updated_df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            # Drop any column starting with # as these are to be ignored and used for comments
            updated_df = updated_df.loc[:, ~updated_df.columns.str.startswith("#")]
        # For any second table onwards, we need to determine extract the column headers by investigating the first row,
        # and then use this to create a completely new data frame with the relevant data.
        else:
            # Assume the first row in data frame represents the new table headers
            table_headers_derived_from_first_row = df.iloc[0, :].values.tolist()
            # Drop columns that are have no value or are empty
            table_headers_derived_from_first_row_excluding_empty_string = [
                value for value in table_headers_derived_from_first_row if value
            ]
            # Determine the column count so that we understand how many headers are to be included
            column_count = len(table_headers_derived_from_first_row_excluding_empty_string)
            # Extract the data in all subsequent rows and represent this as a list of list (one list per row)
            data = df.iloc[1:, 0:column_count].values.tolist()
            # Create a new DF with the relevant values
            updated_df = pd.DataFrame(
                data=data,
                columns=table_headers_derived_from_first_row_excluding_empty_string,
            )
            # Same treatment as the first table, drop any unnamed columns and columns starting with #
            updated_df = updated_df.loc[:, ~updated_df.columns.str.contains("^Unnamed")]
            updated_df = updated_df.loc[:, ~updated_df.columns.str.startswith("#")]

        # Check how many total rows are available in the table
        total_rows_in_table = updated_df.shape[0]
        # Show the table information
        print(f"Table {table_no}:")
        print(updated_df)
        # Ignore the table if it has no rows
        if total_rows_in_table == 0:
            print("Table has no row - ignoring data frame.")
        else:
            table_data_frames.append(updated_df)
            table_no += 1
        print("-------------------------------------------------------------------")

    return table_data_frames


def extract_variables_from_data_frame(data_frame: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    """Extract the variables from the the pandas data frame.

    To make this as extensible as possible, all worksheets within the variables spreadsheet are presented
    as a key within a dictionary with their values representing a list of dicts for each row in the table.

    If the worksheet contains multiple tables, then a unique key will be added for each table, i.e. if the worksheet
    is called 'devices' and it contains two tables, they will appear as:

    devices_table_1
    devices_table_2

    If there is only a single table in the worksheet, then it can be referenced as:

    devices  AND/OR
    devices_table_1

    Args:
        data_frame: the pandas data frame containing the content of a excel spreadsheet

    Returns:
        variables: A dictionary containing the entire output of the excel spreadhseet in a specific format

        Example format:
        {
            # Example of horiztonal table
            "worksheet1_table_1": [
                {"device": "router1", "model": "c8300"},
                {"device": "router2", "model": "c8300"},
            ],
            # Example of vertical table
            "worksheet1_table_2": {
                "ntp_server": "1.1.1.1",
                "aaa_server": "2.2.2.2",
            }
        }
    """
    # The variable dictionary returned as part of this function
    variables: Dict[str, Any] = {}

    # Iterate over all the worksheets stored within the pandas data frame
    for worksheet_name in data_frame:
        print(f"processing worksheet: {worksheet_name}")
        worksheet_df = data_frame[worksheet_name]
        # Specific hardcoded worksheets should only contain a single table - we automatically drop any empty rows
        # to prevent the code from trying to divide the data frames into multiple tables/dfs.
        if worksheet_name in worksheets_with_single_table_allowed:
            worksheet_df.dropna(inplace=True)

        # For all other worksheets, break the data frames into multiple smaller DFs, one for each table
        worksheet_table_data_frames = create_data_frame_for_each_table(data_frame=worksheet_df)
        for index, table_df in enumerate(worksheet_table_data_frames):
            # Generate the automated name to assign as the variable key
            combined_worksheet_and_table_name = f"{worksheet_name}_table_{index+1}"
            total_columns_in_table = len(table_df.axes[1])
            # If the table only contains two columns, assume it's a vertical table and create a k/v dict
            if total_columns_in_table == 2:
                table_dict: Dict[Any, Any] = {}
                first_column, second_column = table_df.iloc[:, 0], table_df.iloc[:, 1]
                for k, v in zip(first_column, second_column):
                    table_dict[k] = v
                # The output is referencible via both the worksheet name and combined name for ease of use
                variables[worksheet_name] = table_dict
                variables[combined_worksheet_and_table_name] = table_dict
            # Otherwise assume it's a horiztonal table and create a list of dicts
            else:
                # The output is referencible via both the worksheet name and combined name for ease of use
                variables[worksheet_name] = table_df.to_dict(orient="records")
                variables[combined_worksheet_and_table_name] = table_df.to_dict(orient="records")
    return variables


def generate_design_diagrams(
    visio_template_path: str,
    save_to_path: str,
    diagram_variable_values: Dict[str, Any],
    diagram_label_values: List[Dict[str, Any]],
) -> bool:
    """Using a template visio file, replace all place holder text and then create a new visio file.

    NOTE: the diagram label values are first rendered, followed by any jinja 2 diagram variable values.

    Args:
        visio_template_path: the full path and filename where the template visio file is located
        save_to_path: the full path and filename where the new visio file will be created
        diagram_variable_values: a key, value dictionary containing the jinja2 variables to be rendered
            Example:
                {
                    "spine1_name": "syd_sw_01",
                    "spine2_name": "syd_sw_02",
                }
        diagram_label_values: a list of dictionaries containing the diagram label + values to update
            Example:
                [
                    {"label": "role", "value": "spine_1", "replacement_text": "syd_sw_01"},
                    {"label": "role", "value": "spine_2", "replacement_text": "syd_sw_02"},
                ]

    Returns:
        design_diagrams_generated: True if the new visio file was created, otherwise False if an error occurs
    """
    design_diagrams_generated: bool = False
    print(f"Reading detailed design template diagram from: {visio_template_path}")
    try:
        with VisioFile(visio_template_path) as vis:
            # Each visio diagram may have one or more pages
            for page in vis.pages:
                print(f" - processing page: {str(page.name)}")
                # The first step is to replace the text for all visio shapes that match label/value combination
                # with the replacement text defined in the variables spareadsheet.
                for update in diagram_label_values:
                    # If a shape contains the matching label and value, then replace the text
                    shape = page.find_shape_by_property_label_value(update["label"], update["value"])
                    if shape:
                        original_text_value = str(shape.text).strip("\n")
                        replacement_text_value = str(update["replacement_text"])
                        print(f"   - replaced value: {original_text_value} - with: {replacement_text_value}")
                        shape.text = update["replacement_text"]
                # After applying the label changes, now remove any shape that has the text 'remove_me', including
                # the removal of any connectors attached to the shape.
                shapes_to_remove = page.find_shapes_by_text("remove_me")
                for shape_to_remove in shapes_to_remove:
                    connected_shapes = shape_to_remove.connected_shapes
                    for connected_shape in connected_shapes:
                        connected_shape.remove()
                    print(f"   - auto deleted shape called 'remove_me' and all it's connected objects")
                    shape_to_remove.remove()
            # If the visio diagram has any jinja2 variables defined, i.e. {{ spine }} and this value is defined
            # within the variables spreadsheet under the 'diagram_variables' tab, then automatically substitute the
            # value as the final step before saving the daigram
            vis.jinja_render_vsdx(context=diagram_variable_values)
            # Export to a new visio file
            vis.save_vsdx(save_to_path)  # save to a new file
            vis.close_vsdx()
            design_diagrams_generated = True
            print(f"Generated new detailed design diagram at: {save_to_path}")
    except Exception as e:
        print(f"Failed to generate new design diagrams at: {save_to_path} - due to: {e}")
    return design_diagrams_generated


def close_and_save_running_visio_application() -> None:
    """Close and saves and visio applications that are running.

    This script relies on win32com which runs visio as a background application to actually export the visio
    file to other image formats such as .png. If the application is already running then the script will fail.

    NOTE: the application will be automatically closed and all outstanding changes will be saved.

    Args:
        N/A

    Returns:
        N/A
    """
    # Visio must be closed otherwise this script will fail. Auto close it now if it's running.
    try:
        visio_running = win32com.client.GetActiveObject("Visio.Application")  # type: ignore[no-untyped-call]
        if visio_running:
            print("Visio was found running - auto closing (changes will be saved).")
            visio_running.Quit(SaveChanges=True)
    # Deliberately swallow the exception - if the application is not running then we don't want it to do anything.
    except Exception:
        pass


def close_and_save_running_word_application() -> None:
    """Close and saves and word applications that are running.

    If the script is run while the old design is open by the user, an error is occured. Therefore we deliberately
    force any word applications to save and close to avoid errors.

    Args:
        N/A

    Returns:
        N/A
    """
    try:
        word_running = win32com.client.GetActiveObject("Word.Application")  # type: ignore[no-untyped-call]
        if word_running:
            print("Word was found running - auto closing (changes will be saved).")
            word_running.Quit(SaveChanges=True)
    # Deliberately swallow the exception - if the application is not running then we don't want it to do anything.
    except Exception:
        pass


def export_visio_diagrams_to_png(visio_diagram_path: str, save_to_path: str) -> bool:
    """Exports all pages in a visio file to .png files matching the page name.

    Args:
        visio_diagram_path: the full path including filename where the visio file is located
        save_to_path: the base path where all the .png files will be saved to

    Returns:
        visio_file_exported_to_ping: True if all the pages were successfully exported to .png otherwise False
    """
    visio_file_exported_to_ping: bool = False
    try:
        visio = win32.gencache.EnsureDispatch("Visio.Application")
    # Sometimes there is a cache issue and visio cannot be opened. The fix is to delete the gen_py directory
    # under the local user folder, see: https://stackoverflow.com/questions/70121855/open-visio-using-python
    except AttributeError as e:
        user_tmp_dir = os.environ["USERPROFILE"]
        gen_py_tmp_dir = os.path.join(user_tmp_dir, "AppData", "Local", "Temp", "gen_py")
        print("Error opening visio due to cache issue.")
        print(f"Attempting to delete the cache folder at: {gen_py_tmp_dir}")
        shutil.rmtree(gen_py_tmp_dir)
        print(" - Cache has been deleted. Re-run this script and it should work the next time.  time.")
        sys.exit(0)
    # Hide the window when it opens
    visio.Visible = True

    try:
        print(f"Visio images will be temporarily exported from: {visio_diagram_path}")
        doc = visio.Documents.Open(visio_diagram_path)
        for page in doc.Pages:
            page_name = str(page)
            # Use the visio page name as the file name
            image_filename = page_name + ".png"
            image_export_path = os.path.join(save_to_path, image_filename)
            page.Export(image_export_path)
            print(f"  - exported '{str(page_name)}' page to: {image_export_path}")
        doc.Close()
        visio_file_exported_to_ping = True
    except Exception as e:
        print(f"UNABLE TO Export VISO FILE to PNG DUE TO : {e}")
    visio.Quit()
    return visio_file_exported_to_ping


def generate_design_document(doc_template: DocxTemplate, save_to_path: str, variables: Dict[str, Any]) -> bool:
    """Generates a new word document based on the template after the values have been rendered.

    Args:
        doc_template: instance of the word document template that has already been loaded into memory
        save_to_path: the full path and filename where the rendered design document will be saved to
        variables: a dictionary containing the variables that are available to the document for jinja2 rendering
            Example:
                {
                    "worksheet1": [{"column1": "example", "column2": "example"}],
                    "worksheet1_table_1": [{"column1": "example", "column2": "example"}],
                    "worksheet1_table_2": {"k1": "v1", "k2": "v2"},
                    "worksheet2_table_1": {"k1": "v1", "k2": "v2"},
                }

    Returns:
        design_document_generated: True if the design document was generated without errors, otherwise False
    """
    design_document_generated: bool = False
    # Replace jinja 2 formatting in the word documenta nd save it to a new file
    try:
        doc_template.render(variables)
        doc_template.save(save_to_path)
        print(f"New design documented generated and saved to: '{save_to_path}'")
        design_document_generated = True
    except PermissionError as e:
        print(f"Unable to generate new design document at: {save_to_path} - due to: {e}")
    except UndefinedError as e:
        print(f"ERROR: The template is referencing a worksheet or table that doesn't exists: {e}")
        sys.exit(0)
    return design_document_generated


def add_images_to_template_variables(
    doc_template: DocxTemplate, variables: Dict[str, Any], image_path: str
) -> Dict[str, Any]:
    """Extend the template variables to include all *.png files that are located within a specific path.

    In order to reference images as jinja2 variables within the template document, we need to add them using
    a special InlineImage object.

    Args:
        doc_template: instance of the word template that has already been loaded into memory
        image_path: the path to search for *.png files
        variables: the pandas data frame after it has been exported into a dictionary
            Example:
                {
                    "worksheet1": [{"column1": "example", "column2": "example"}],
                    "worksheet2": [{"column1": "example", "column2": "example"}],
                }

    Returns:
        variables: updated variables dictionary with a new 'images' keyword
            Example:
                {
                    "worksheet1": [{"column1": "example", "column2": "example"}],
                    "worksheet2": [{"column1": "example", "column2": "example"}],
                    "images": {
                        "image1.png": InlineImage(),
                        "image2.png": InlineImage(),
                    }
                }
    """
    variables["images"] = {}
    # Iterate over the speific image path and locate all .png files within the directory
    path = pathlib.Path(image_path)
    for f in path.iterdir():
        if not str(f.name).endswith(".png"):
            continue
        image_filename = str(f.name)
        image_filename_without_extension = str(f.name).replace(".png", "")
        # Add the image object to the variables dictionary
        variables["images"][image_filename_without_extension] = InlineImage(doc_template, image_filename)
    return variables


def remove_png_files_from_path(image_path: str) -> bool:
    """Removes *.png files from a specific path.

    This is used as a clean up operation to ensure that any temporary *.png files are removed after the script
    is complete.

    Args:
        image_path: the full path to the directory to remove .png files from

    Returns:
        all_png_files_removed: returns True if all .png files were successfully removed otherwise False
    """
    all_png_files_removed: bool = True
    path = pathlib.Path(image_path)
    for f in path.iterdir():
        if not str(f.name).endswith(".png"):
            continue
        try:
            os.remove(f)
            print(f"Temporary file: {f} has been removed.")
        except Exception as e:
            print(f"Unable to remove {f} - due to: {e}")
            all_png_files_removed = False
    return all_png_files_removed

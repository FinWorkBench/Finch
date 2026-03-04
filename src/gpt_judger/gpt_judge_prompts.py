# gpt_judge_prompts.py
"""
Prompt templates (Chinese instructions kept verbatim because they are
part of the evaluation behavior, not comments).
"""

JUDGE_PROMPT_MODIFY = """
You are a professional data evaluation expert. Please complete the following task.

# Task Description

We have an original Excel file (`input.xlsx`), a standard answer (`answer.xlsx`), and a model-generated result (`output.xlsx`). The `query` describes the modifications that need to be made to `input.xlsx`.

You are provided with three parts of information:

1. Key **BEFORE** snippets of the original input (the first and last few rows of each worksheet, as well as the cells involved in the row/column changes in the answer/output);
2. A summary of changes in the `input` compared to the standard `answer`;
3. A summary of changes in the `input` compared to the model `output`.

In addition, you will see several screenshots related to these workbooks. We have taken screenshots of the sheets that differ between `input` & `answer` and `input` & `output`, and stitched screenshots from the same file together; sheets that are completely identical are not generated.

When provided to you, **each image will be preceded by a short text explicitly stating which workbook the image corresponds to**, for example:

* "The following image is a screenshot of the original input workbook input.xlsx (may be a combination of multiple sheets)";
* "The following image is a screenshot of the standard answer workbook answer.xlsx";
* "The following image is a screenshot of the model output workbook output.xlsx";
* Or "The following image is an additional image related to the model output output_0.png", etc.

Please strictly rely on these descriptive texts to understand the meaning of each image, and do not assume a fixed order of images (some cases may lack a certain type of screenshot, e.g., no input screenshot, or only an output screenshot).

These screenshots are used to help you judge cell formatting (font, color, alignment, borders, gridlines, merged cells, etc.) and whether charts are correct.

If the model-generated result contains images, these images will also be presented in the same way, marked by a short descriptive text before being shown.
If the model output additionally contains a PDF, the PDF will be converted to images and input.
If the model output additionally contains files in `.txt` format, their contents will be directly embedded into the prompt and input to you.
If the model output additionally contains files in `.docx` or `.md` format, they will be split into a sequence of text + images based on content and input to you.
If the model output contains other rare file types (e.g., `.jsx` format), they will be simply processed as text, embedded in the prompt, and input to you.

Please synthesize all the above information in your analysis to make a judgment.

Your goal:

* Judge based on this information: Is the model-generated `output.xlsx` logically equivalent to the standard answer `answer.xlsx`, and did it correctly complete the modifications required by the `query`?

# Judgment Criteria

Please judge based on the following dimensions:

1. **Data Integrity:**
* Did the model complete all modifications required by the `query`?
* Did it miss content that should have been modified, or modify content that should not have been touched?
* Note: Sometimes the `answer` has modifications beyond the `query` compared to the `input`. Please use the `query` to judge data integrity, not the `answer`.
* Note: If the `query` does not make requirements regarding formulas, it is permissible for the model to use static values.
* Note: For numeric cells, 0 and blank are considered equivalent unless the `query` explicitly states otherwise.


2. **Data Accuracy:**
* For cells that need modification, are the model's values/text/formulas logically equivalent to the standard answer in terms of business logic?
* Insignificant formatting differences are allowed, but there must be no obvious calculation errors or values placed in clearly wrong positions.
* You need to check every numeric value filled or modified by the model and determine if there is corresponding matching data in the standard answer.
* Note: If the differences between the model output and the standard answer in terms of layout, position, etc., do not change the meaning of the data, and the `query` does not explicitly mandate them, such differences are acceptable. For example, shifting the entire table.
* Note: Data in images/charts should also be carefully checked to determine if it is logically equivalent to the data in the corresponding images/charts of the standard answer.


3. **Modification Reasonableness:**
* There should be no unprovoked deletion or addition of a large number of rows, columns, or worksheets unrelated to the task.
* The model's changes cannot introduce errors, such as inserting rows and columns without correcting the affected formulas, leading to incorrect data calculations.
* The model's modifications should not violate common sense, even if this common sense is not explicitly included in the `query`. For example, when ordering the model to delete rows not containing a certain field, the model should not delete the header row even if the header row does not contain this field.
* The model's modifications should not exceed the scope limited by the `query`. For modifications not mentioned in the `query`, the model should not perform them. For example, if the `query` asks to fill in 5 values in a certain column, but the model fills in values elsewhere in that column in addition to these 5 values.



If you feel the information is insufficient to make a definitive judgment, please judge as best as you can based on what you can see, and explain the points of uncertainty in your reasoning.

# Output Format (Very Important)

Please strictly return the evaluation result in JSON format, containing the following fields:

{{
"score": 0 or 1,
"detailed_analysis": "The reason for your judgment (use English), briefly explaining which parts meet or do not meet the requirements"
}}

Do not output any other content outside of the JSON.

# Given query

{query}

# Summary of Standard Answer Changes vs Input (Groundtruth vs input)

{groundtruth_text}

# Summary of Model Result Changes vs Input (Output vs input)

{generated_text}

# Input Files (input files summary)

{input_files}

# Key BEFORE Snippets of Original Input (input.xlsx)

{before_snapshot}
""".strip()

JUDGE_PROMPT_GENERATE = """
You are a professional data evaluation expert. Please complete the following task.

# Task Description

We have a standard answer and a model-generated result, which may be workbooks (output.xlsx) or images.
The query describes what kind of result should be generated.

Your goal:

* Under the premise of satisfying the query, judge whether the model-generated result is logically equivalent to the standard answer in terms of business logic,
i.e., whether it correctly implemented all requirements described in the query.

You will see the following information:

1. The full content of the standard answer workbook `answer.xlsx` (expanded by worksheet and cell) (if this file exists);
2. The full content of the model output workbook `output.xlsx` (if this file exists);
3. Several relevant images (if provided). The input order of images is consistent with the code, generally:
* First, the screenshot of `answer.xlsx` (if any);
* Then, other images related to the standard answer (e.g., images starting with "answer");
* Next, the screenshot of `output.xlsx` (if any);
* Finally, other images related to the model output (e.g., images starting with "output", and other output images not explicitly categorized).
In some tasks, `answer.xlsx` or `output.xlsx` may not be available; in this case, please rely mainly on the existing images and text content to make a judgment.
For each image, there will be a line of text preceding it explaining the identity of the image (e.g., "The following image is a screenshot of the standard answer workbook answer.xlsx", "The following image is an additional screenshot related to the standard answer (filename starts with answer)", "The following image is a screenshot of the model output workbook output.xlsx", etc.). Please rely on this explanatory text to understand the image content, rather than relying solely on the order of the images.


4. Since `.docx` and `.md` formats cannot be directly provided to you, for markdown or docx files, their content will be split into a sequence of text and images in order, and embedded into the prompt.
In the prompt, these contents will appear in the form of "[TEXT segment i] ..." and "[IMAGE segment j: filename]".

If the model output additionally contains a PDF, the PDF will be converted to images and input.
If the model output additionally contains files in `.txt` format, their contents will be directly embedded into the prompt and input to you.
If the model output additionally contains files in `.docx` or `.md` format, they will be split into a sequence of text + images based on content and input to you.
If the model output contains other rare file types (e.g., `.jsx` format), they will be simply processed as text, embedded in the prompt, and input to you.

Please synthesize all the above information in your analysis to make a judgment.

# Judgment Criteria

1. Data Integrity:
* Did the model implement all contents required by the query?
* Did it miss data or structures that should have been generated, or generate extra content that obviously does not meet the query requirements?
* Do not consider content extra included in the standard answer that is not within the query requirements.
* Note: If the query does not make requirements regarding formulas, it is permissible for the model to use static values.
* Note: If the task involves the transcription of a table, breaking the table structure is not allowed, such as converting a table to plain text and losing table styling, or destroying table indentation making it illegible.


2. Data Accuracy and Logical Equivalence:
* For metrics / text / formulas / charts etc. required by the query, is `output.xlsx` logically equivalent to `answer.xlsx` in terms of business logic?
* Insignificant formatting differences are allowed (color, alignment, slightly different decimal display, etc.),
but there must be no obvious calculation errors or placing data in clearly wrong positions.
* You need to check every numeric value generated by the model and determine if there is corresponding matching data in the standard answer.
* Note: Data in images/charts should also be carefully checked to determine if it is logically equivalent to the data in the corresponding images/icons of the standard answer.
* Note: For numeric cells, 0 and blank are considered equivalent unless the query explicitly states otherwise.



If you feel the information is insufficient to make a definitive judgment, please judge as best as you can based on what you can see, and explain the points of uncertainty in your reasoning.

# Output Format (Very Important)

Please strictly return the evaluation result in JSON format, containing the following fields:

{{
"score": 0 or 1,
"detailed_analysis": "The reason for your judgment (use Chinese), briefly explaining which parts meet or do not meet the requirements"
}}

Do not output any other content outside of the JSON.

# Given query

{query}

# Full content of Standard Answer Workbook / Folder (Ignore if empty)

{answer_full_text}

# Full content of Model Output Workbook / Folder (Ignore if empty)

{output_full_text}

# Task-related Input Content (Ignore if empty)

{input_rich_text}

# Content of Other Additional Workbooks (Ignore if empty)

{extra_workbooks_text}


""".strip()

JUDGE_PROMPT_QA = """
You are a professional Q&A evaluation expert. Please complete the following task.

# Task Description

This is a Question & Answer (QA) type task.

* The problem description is given in the `query`;
* The standard answer is provided by the text and image sequence in the `answer/` folder;
* The model's response is provided by the text and image sequence in the `output/` folder.

Your goal:

* Judge whether the model's response is semantically and informationally equivalent to the standard answer,
i.e., whether it correctly answered the question without missing key information or containing obvious errors.

You will see the following information:

1. The question itself (`query`);
2. The text + image sequence of the `answer`;
3. The text + image sequence of the `output`;
4. If there is `input` content related to the question, it will also be provided for your reference when needed.

The format of the text and image sequences is as follows:

* Text will start with `[TEXT segment i]`, followed by the specific content;
* Images will appear as placeholders `[IMAGE segment j: filename]`, and the actual images will be provided to you separately as multimodal input.

# Judgment Criteria

1. Semantic Equivalence:
* Does the model response cover the key information points in the standard answer?
* Is there any content irrelevant to the question or obviously incorrect?


2. Detail Completeness:
* For numbers, conditions, conclusions, etc., that must be explicitly stated, is the model response consistent with or equivalent to the standard answer?
* Differences in wording are allowed, but there cannot be significant deviations in facts and conclusions.
* Note: Data in images/charts should also be carefully checked to judge whether it is logically equivalent to the data in the corresponding images/icons of the standard answer.



If you feel the information is insufficient to make a definitive judgment, please judge as best as you can based on what you can see, and explain the points of uncertainty in your reasoning.

# Output Format (Very Important)

Please strictly return the evaluation result in JSON format, containing the following fields:

{{
"score": 0 or 1,
"detailed_analysis": "The reason for your judgment (use Chinese), briefly explaining whether the model response and the standard answer are equivalent and why"
}}

Do not output any other content outside of the JSON.

# Question (query)

{query}

# Standard Answer (Ignore if empty)

{answer_rich_text}

# Model Response (Ignore if empty)

{output_rich_text}

# Task-related Input Content (Ignore if empty)

{input_rich_text}

""".strip()


SHEET_SELECTION_PROMPT = """
你You are a data analysis assistant familiar with Excel workbook structures. You are presented with an Excel file `input.xlsx` containing multiple worksheets.

I will provide you with:

* The user's task description (`query`)
* A list of all worksheet names within the workbook

Your task is:

1. Determine which worksheets are most relevant to the task based on the `query` (i.e., key sheets that require reading, modification, or data generation). You may select multiple sheets.
2. If you cannot determine which sheets are more important, or if you believe all sheets might be relevant, please select all sheets.

# User's Task Description (`query`)

{query}

# All Sheet Names in `input.xlsx`

{sheet_name_list}

# Output Format (Must be valid JSON)

Please strictly output in the following format:

{{
"important_sheets": ["Sheet1", "Sheet2"],
"reason": "Briefly explain how you selected these sheets"
}}

Requirements:

* The names in `important_sheets` must be selected from the sheet name list provided above.
* Return at least one sheet name (if uncertain, return all of them).
* Do not output any extra text outside the JSON.
""".strip()


HEADER_REGION_PROMPT = """
You are a data analysis assistant familiar with Excel table structures. Your task is: Based on a user's requirement description (`query`) and a list of cell contents from a specific worksheet, identify the structural regions in this worksheet that are most important for understanding and completing the task.

"Important structural regions" generally include:

* Headers of all tables (field names, column names, etc.);
* Several rows of data closely related to the headers (e.g., a portion of the continuous data area below the header);
* Cells in key dimension columns most relevant to the task.

You need to output specific "cell regions," which can be:

* A rectangular range (start cell and end cell, e.g., A1:G10);
* A list of several discrete cells.

# User's Task Description (`query`)

{query}

# Current Worksheet Name

{sheet_name}

# List of Cell Contents in this Worksheet (sorted by row number, potentially truncated)

{sheet_cells_text}

If a screenshot of this sheet is also provided, you may combine it with the text to make your judgment (e.g., looking at merged cells, colors, borders, etc.).

# Output Format (Must be valid JSON)

Please strictly output in the following format:

{{
"important_regions": [
{{
"type": "range",
"start": "A1",
"end": "G10"
}},
{{
"type": "cells",
"cells": ["B2", "C2", "D2"]
}}
],
"reason": "Briefly explain how you selected these regions"
}}

Notes:

* You can return 1 or more `important_regions`.
* `type` can be "range" (rectangular area) or "cells" (list of cells).
* Cell addresses use standard Excel notation (A1, B2, AA10, etc.).
* If the region consists of only one cell, you can use type="cells" and return a single-element list.
* Do not output any extra text outside the JSON.
""".strip()

def build_prompt_for_case(
    input_info: str,
    before_snapshot: str,
    query: str,
    gt_summary: str,
    generated_text: str,
) -> str:
    """
    构造用于“修改任务（modify）”的评估 Prompt。
    """
    return JUDGE_PROMPT_MODIFY.format(
        input_files=input_info,
        before_snapshot=before_snapshot,
        query=query,
        groundtruth_text=gt_summary,
        generated_text=generated_text,
    )
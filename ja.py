# Make sure to run `datamodel-codegen --input cv.json --output prg/__generated__/cv_model.py`
# when cv.json changes and to update the selectors if needed.

# This file takes care of shifting to the right if the left column is filled up.
# Run it first before trying to fix a non-issue lol

import shutil
import subprocess
from pypdf import PdfReader, PdfWriter
from spire.xls import *
from spire.xls.common import *
from prg.loadCV import loadCV
from datetime import date

# read cv.json
data = loadCV("cv.json")


# load japanese cv template
wb = Workbook()
wb.LoadFromFile("./ja.template.xlsx")
sheet = wb.Worksheets[0]

# Set the current date
today = date.today()
today = today.strftime("%Y年%m月%d日")
sheet.Range["E3"].Text = today

# Set the name
sheet.Range["B5"].Text = data.Profile.Furigana
sheet.Range["B7"].Text = data.Profile.Name.ja

# Set the nationality
sheet.Range["B13"].Text = data.Profile.Nationality.ja

# Set gender
sheet.Range["L13"].Text = data.Profile.Gender.ja

# Set the birthdate
year = data.Profile.Birthyear
month = data.Profile.Birthmonth
day = data.Profile.Birthday
sheet.Range["G13"].Text = f"{year}年{month}月{day}日"

# Set the address
sheet.Range["B15"].Text = data.Profile.Address.Furigana
sheet.Range["B17"].Text = f"〒{data.Profile.Address.Zip}"
sheet.Range["B19"].Text = data.Profile.Address.Address

# Set the phone number
sheet.Range["K16"].Text = data.Profile.Phone

# Set the email
sheet.Range["K19"].Text = data.Profile.Email

# Fill in education title
base = 30
sheet.Range[f"C{base}"].Text = "学歴"

# Center the title
sheet.Range[f"C{base}"].HorizontalAlignment = HorizontalAlignType.Center

# Fill in education history
base += 2
base_letters = ["A", "B", "C"]
for i, item in enumerate(data.Education):
    row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + row_offset}")
    sheet.Range[f"{base_letters[1]}{base + row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + row_offset}")

    # Set the description
    for i, detail in enumerate(item.Details):

        if base_letters == ["O", "P", "Q"] and base + row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        description = f"{detail.Description.ja}"
        sheet.Range[f"{base_letters[2]}{base + row_offset}"].Text = description

        if base_letters == ["O", "P", "Q"] and base + row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column in the education. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        # printing pretty messages
        if len(description) > 7:
            print_description = description[:7] + "..."
        ansi_red_coordinates = f"\033[91m{base_letters[2]}{base + row_offset}\033[0m"
        print(f"Putting {print_description} in {ansi_red_coordinates}")

        # Align description to the left
        sheet.Range[f"{base_letters[2]}{base + row_offset}"].HorizontalAlignment = (
            HorizontalAlignType.Left
        )
        row_offset += 2

        # Move to column on the right if the end of the left one is reached
        if base + row_offset >= 62:
            # ansi blue "education history" message
            shift_during = "\033[94meducation history\033[0m"

            # ansi green "Moving rest of the education history to the right" message
            print(
                f"\033[92mMoving rest of the {shift_during} \033[92mto the right\033[0m"
            )

            # shift to the right
            base_letters = ["O", "P", "Q"]
            base = 5
            row_offset = 0

# Fill in work title
work_title_base = base + row_offset + 2
sheet.Range[f"{base_letters[2]}{work_title_base}"].Text = "職歴"

# Center the title
sheet.Range[f"{base_letters[2]}{work_title_base}"].HorizontalAlignment = (
    HorizontalAlignType.Center
)

# Fill in work history
base = work_title_base + 2
base_letters = base_letters
for i, item in enumerate(data.Work):
    row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + row_offset}")
    sheet.Range[f"{base_letters[1]}{base + row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + row_offset}")

    # Set the description
    for i, detail in enumerate(item.Details):
        if base_letters == ["O", "P", "Q"] and base + row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column in the work history. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        description = f"{detail.Description.ja}"
        sheet.Range[f"{base_letters[2]}{base + row_offset}"].Text = description

        # printing pretty messages
        if len(description) > 7:
            print_description = description[:7] + "..."
        ansi_red_coordinates = f"\033[91m{base_letters[2]}{base + row_offset}\033[0m"
        print(f"Putting {print_description} in {ansi_red_coordinates}")

        # Align description to the left
        sheet.Range[f"{base_letters[2]}{base + row_offset}"].HorizontalAlignment = (
            HorizontalAlignType.Left
        )
        row_offset += 2

        # Move to column on the right if the end of the left one is reached
        if base + row_offset >= 62:
            # ansi blue "work history" message
            shift_during = "\033[94mwork history\033[0m"

            # ansi green "Moving rest of the work history to the right" message
            print(
                f"\033[92mMoving rest of the {shift_during} \033[92mto the right\033[0m"
            )

            # shift to the right
            base_letters = ["O", "P", "Q"]
            base = 5
            row_offset = 0

# Put Projects into special qualifications
base = 24
base_letters = ["O", "P", "Q"]
for i, item in enumerate(data.Projects):
    row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + row_offset}")
    sheet.Range[f"{base_letters[1]}{base + row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + row_offset}")

    # Set the description

    if base_letters == ["O", "P", "Q"] and base + row_offset >= 38:
        # ansi red "error" message
        error = "\033[91mYou have reached the end of the right column in the project list. Please adjust the code to handle this case.\033[0m"
        print(error)
        break

    description = f"{item.Title} - {item.Details.ja}"
    sheet.Range[f"{base_letters[2]}{base + row_offset}"].Text = description

    # printing pretty messages
    if len(description) > 7:
        print_description = description[:7] + "..."
    ansi_red_coordinates = f"\033[91m{base_letters[2]}{base + row_offset}\033[0m"
    print(f"Putting {print_description} in {ansi_red_coordinates}")

    # Align description to the left
    sheet.Range[f"{base_letters[2]}{base + row_offset}"].HorizontalAlignment = (
        HorizontalAlignType.Left
    )
    row_offset += 2


# gather design skills
design_tech = data.Skills.Design.Technology

# gather programming skills
programming_tech = data.Skills.Programming.Technology

# gather language skills
language_tech = data.Skills.Languages.Technology

# gather hobbies
hobbies_tech = data.Skills.Hobbies.Technology

# gather other skills
skills_tech = data.Skills.Skills.Technology

# combine all skills
skills = f"プログラミング: {programming_tech}\nデザイン: {design_tech}\n他のスキル: {skills_tech}\n言語: {language_tech}\n趣味: {hobbies_tech}"

# Fill in skills
sheet.Range["O40"].Text = skills

# Fill in dependents
sheet.Range["X43"].Text = f"{data.Profile.Dependents}人"
sheet.Range["X43"].HorizontalAlignment = HorizontalAlignType.Center

# Fill in marital status
sheet.Range["X46"].Text = data.Profile.MaritalStatus.ja

# Fill in alimony payment information
sheet.Range["Z46"].Text = data.Profile.AlimonyPayments.ja

# Fill in expectations
workplace_expectations = data.Expectations.Workplace.ja
salary_expectations = data.Expectations.Salary.ja
sheet.Range["O51"].Text = f"職場: {workplace_expectations}\n給与: {salary_expectations}"

# save japanese cv
wb.SaveToFile("./out/ja.cv.xlsx", FileFormat.Version2016)

# green ansi success message for xlsx file
xlsx_success_message = "\033[92mSuccessfully generated ja.cv.xlsx\033[0m"
print(xlsx_success_message)

# Run LibreOffice in headless mode to convert the file
subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "./out/ja.cv.xlsx"], check=True
)

# Define new location
soffice_output = "ja.cv.pdf"
destination_path = os.path.join("./out", os.path.basename("ja.cv.pdf"))

# Move the file to the out folder
shutil.move(soffice_output, destination_path)

# crop to first page
reader = PdfReader(destination_path)
writer = PdfWriter()

# Add only the first page to the new PDF
writer.add_page(reader.pages[0])

# Save the new PDF
with open(destination_path, "wb") as output_file:
    writer.write(output_file)

# green ansi success message for pdf file
pdf_success_message = "\033[92mSuccessfully generated ja.cv.pdf\033[0m"
print(pdf_success_message)

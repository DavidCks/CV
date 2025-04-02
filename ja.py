# Make sure to run `datamodel-codegen --input cv.json --output prg/__generated__/cv_model.py`
# when cv.json changes and to update the selectors if needed.

# This file takes care of shifting to the right if the left column is filled up.
# Run it first before trying to fix a non-issue lol

############
# Contents #
############
# ctrl+f the full line to go to it's place

# -------- #
# Contents #
#  Resume  #
#    CV    #
# -------- #

import shutil
import subprocess
from pypdf import PdfReader, PdfWriter
from spire.xls import *
from spire.xls.common import *
from prg.loadCV import loadCV
from datetime import date

# read cv.json
data = loadCV("cv.json")

############
#  Resume  #
############

# load japanese resume template
wb = Workbook()
wb.LoadFromFile("./ja.resume.template.xlsx")
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
    work_row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + work_row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[1]}{base + work_row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")

    # Set the description
    for i, detail in enumerate(item.Details):

        if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        description = f"{detail.Description.ja}"
        sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = description

        if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column in the education. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        # printing pretty messages
        if len(description) > 7:
            print_description = description[:7] + "..."
        ansi_red_coordinates = (
            f"\033[91m{base_letters[2]}{base + work_row_offset}\033[0m"
        )
        print(f"Putting {print_description} in {ansi_red_coordinates}")

        # Align description to the left
        sheet.Range[
            f"{base_letters[2]}{base + work_row_offset}"
        ].HorizontalAlignment = HorizontalAlignType.Left
        work_row_offset += 2

        # Move to column on the right if the end of the left one is reached
        if base + work_row_offset >= 62:
            # ansi blue "education history" message
            shift_during = "\033[94meducation history\033[0m"

            # ansi green "Moving rest of the education history to the right" message
            print(
                f"\033[92mMoving rest of the {shift_during} \033[92mto the right\033[0m"
            )

            # shift to the right
            base_letters = ["O", "P", "Q"]
            base = 5
            work_row_offset = 0

# Fill in work title
work_title_base = base + work_row_offset + 2
sheet.Range[f"{base_letters[2]}{work_title_base}"].Text = "職歴"

# Center the title
sheet.Range[f"{base_letters[2]}{work_title_base}"].HorizontalAlignment = (
    HorizontalAlignType.Center
)

# Fill in work history
base = work_title_base + 2
base_letters = base_letters
for i, item in enumerate(data.Work):
    work_row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + work_row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[1]}{base + work_row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")

    # Set the description
    for i, detail in enumerate(item.Details):
        if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 21:
            # ansi red "error" message
            error = "\033[91mYou have reached the end of the right column in the work history. Please adjust the code to handle this case.\033[0m"
            print(error)
            break

        description = f"{detail.Description.ja}"
        sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = description

        # printing pretty messages
        if len(description) > 7:
            print_description = description[:7] + "..."
        ansi_red_coordinates = (
            f"\033[91m{base_letters[2]}{base + work_row_offset}\033[0m"
        )
        print(f"Putting {print_description} in {ansi_red_coordinates}")

        # Align description to the left
        sheet.Range[
            f"{base_letters[2]}{base + work_row_offset}"
        ].HorizontalAlignment = HorizontalAlignType.Left
        work_row_offset += 2

        # Move to column on the right if the end of the left one is reached
        if base + work_row_offset >= 62:
            # ansi blue "work history" message
            shift_during = "\033[94mwork history\033[0m"

            # ansi green "Moving rest of the work history to the right" message
            print(
                f"\033[92mMoving rest of the {shift_during} \033[92mto the right\033[0m"
            )

            # shift to the right
            base_letters = ["O", "P", "Q"]
            base = 5
            work_row_offset = 0

# Put Projects into special qualifications
base = 24
base_letters = ["O", "P", "Q"]
for i, item in enumerate(data.Projects):
    work_row_offset = i * 2

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + work_row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[1]}{base + work_row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")

    # Set the description

    if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 38:
        # ansi red "error" message
        error = "\033[91mYou have reached the end of the right column in the project list. Please adjust the code to handle this case.\033[0m"
        print(error)
        break

    description = f"{item.Title} - {item.Details.ja}"
    sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = description

    # printing pretty messages
    if len(description) > 7:
        print_description = description[:7] + "..."
    ansi_red_coordinates = f"\033[91m{base_letters[2]}{base + work_row_offset}\033[0m"
    print(f"Putting {print_description} in {ansi_red_coordinates}")

    # Align description to the left
    sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].HorizontalAlignment = (
        HorizontalAlignType.Left
    )
    work_row_offset += 2


# gather design skills
design_tech = data.Skills.Design.Technology.ja

# gather programming skills
programming_tech = data.Skills.Programming.Technology.ja

# gather language skills
language_tech = data.Skills.Languages.Technology.ja

# gather hobbies
hobbies_tech = data.Skills.Hobbies.Technology.ja

# gather other skills
skills_tech = data.Skills.Skills.Technology.ja

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
wb.SaveToFile("./out/ja.resume.xlsx", FileFormat.Version2016)

# green ansi success message for xlsx file
xlsx_success_message = "\033[92mSuccessfully generated ja.resume.xlsx\033[0m"
print(xlsx_success_message)

# Run LibreOffice in headless mode to convert the file
subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "./out/ja.resume.xlsx"], check=True
)

# Define new location
soffice_output = "ja.resume.pdf"
destination_path = os.path.join("./out", os.path.basename("ja.resume.pdf"))

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
pdf_success_message = "\033[92mSuccessfully generated ja.resume.pdf\033[0m"
print(pdf_success_message)

############
#    CV    #
############

print("Generating CV...")

# load japanese cv template
wb = Workbook()
wb.LoadFromFile("./ja.cv.template.xlsx")
sheet = wb.Worksheets[0]

# Set the current date
today = date.today()
today = today.strftime("%Y年%m月%d日")
sheet.Range["G2"].Text = today

# Set the name
sheet.Range["G3"].Text = f"氏名　{data.Profile.Name.ja}"


# Fill in work details
work_base_row = 5
work_row_offset = 0
work_template_row = sheet.Rows[work_base_row]

for i, item in enumerate(data.Work):
    # Copy row for the title
    insertIndex = work_base_row + work_row_offset + 2
    sheet.InsertRow(insertIndex)
    sheet.CopyRow(work_template_row, sheet, insertIndex, CopyRangeOptions.All)

    # Insert title for the new row
    title = f"＜{item.Title.ja}＞"
    sheet.Range[f"A{insertIndex}"].Text = title
    sheet.Range[f"A{insertIndex}"].Style.Font.IsBold = True
    work_row_offset += 1

    for j, detail in enumerate(item.Details):
        # Copy row for the bullet
        insertIndex = work_base_row + work_row_offset + 2
        sheet.InsertRow(insertIndex)
        sheet.CopyRow(work_template_row, sheet, insertIndex, CopyRangeOptions.All)

        # Insert title for the new row
        bullet = f"・{detail.Description.ja}"
        sheet.Range[f"A{insertIndex}"].Text = bullet
        sheet.Range[f"A{insertIndex}"].Style.WrapText = True

        work_row_offset += 1

    # Put empty row
    sheet.InsertRow(insertIndex + 1)
    sheet.CopyRow(work_template_row, sheet, insertIndex + 1, CopyRangeOptions.All)
    sheet.Range[f"A{insertIndex + 1}"].Text = ""

    work_row_offset += 1


# Delete template row
sheet.DeleteRow(work_base_row + 1)


# Fill in skill details
skill_base_row = work_base_row + work_row_offset + 1
skill_row_offset = 0
skill_template_row = sheet.Rows[skill_base_row]

for i, item in enumerate(data.Skills):
    # Copy row for the title
    insertIndex = skill_base_row + skill_row_offset + 2
    sheet.InsertRow(insertIndex)
    sheet.InsertRow(insertIndex + 1)
    sheet.InsertRow(insertIndex + 2)
    sheet.InsertRow(insertIndex + 3)
    sheet.InsertRow(insertIndex + 4)
    sheet.CopyRow(skill_template_row, sheet, insertIndex, CopyRangeOptions.All)
    sheet.CopyRow(skill_template_row, sheet, insertIndex + 1, CopyRangeOptions.All)
    sheet.CopyRow(skill_template_row, sheet, insertIndex + 2, CopyRangeOptions.All)
    sheet.CopyRow(skill_template_row, sheet, insertIndex + 3, CopyRangeOptions.All)
    sheet.CopyRow(skill_template_row, sheet, insertIndex + 4, CopyRangeOptions.All)

    # Insert title for the new row
    titles_ja = {
        "Design": "デザイン",
        "Programming": "プログラミング",
        "Languages": "言語",
        "Hobbies": "趣味",
        "Skills": "他のスキル",
    }
    # Insert Skill title
    title = f"＜{titles_ja[item[0]]}＞"
    sheet.Range[f"A{insertIndex}"].Text = title
    sheet.Range[f"A{insertIndex}"].Style.Font.IsBold = True
    sheet.Range[f"A{insertIndex}"].Style.WrapText = True

    # Insert Skill details
    details = f"{item[1].Details.ja}"
    sheet.Range[f"A{insertIndex + 1}"].Text = details
    sheet.Range[f"A{insertIndex + 1}"].Style.WrapText = True

    # Add spacing between description and technology
    sheet.Range[f"A{insertIndex + 2}"].Text = ""
    sheet.Range[f"A{insertIndex + 2}"].Style.Font.Size = 6
    sheet.Range[f"A{insertIndex + 2}"].Style.VerticalAlignment = (
        VerticalAlignType.Bottom
    )

    # Insert Skill technology
    technology = f"{item[1].Technology.ja}"
    sheet.Range[f"A{insertIndex + 3}"].Text = technology
    sheet.Range[f"A{insertIndex + 3}"].Style.WrapText = True
    sheet.Range[f"A{insertIndex + 3}"].Style.Color = Color.get_LightGray()
    leftBorder = sheet.Range[f"A{insertIndex + 3}"].Borders[BordersLineType.EdgeLeft]
    leftBorder.LineStyle = LineStyleType.Thick
    leftBorder.Color = Color.get_DarkGray()

    # Clear empty template row
    sheet.Range[f"A{insertIndex + 4}"].Text = ""

    skill_row_offset += 5

# Delete template row
sheet.DeleteRow(skill_base_row + 1)

# Fill in work details
projects_base_row = skill_base_row + skill_row_offset + 1
projects_row_offset = 0
projects_template_row = sheet.Rows[projects_base_row]

for i, item in enumerate(data.Projects):
    # Copy row for the title
    insertIndex = projects_base_row + projects_row_offset + 2
    sheet.InsertRow(insertIndex)
    sheet.CopyRow(projects_template_row, sheet, insertIndex, CopyRangeOptions.All)

    # Insert title for the new row
    title = f"＜{item.Title}＞"
    if item.Users > 50:
        title += f"　ー　{item.Users}名"
    sheet.Range[f"A{insertIndex}"].Text = title
    sheet.Range[f"A{insertIndex}"].Style.Font.IsBold = True
    projects_row_offset += 1

    # Put new row for description
    insertIndex = projects_base_row + projects_row_offset + 2
    sheet.InsertRow(insertIndex)
    sheet.CopyRow(projects_template_row, sheet, insertIndex, CopyRangeOptions.All)

    # Insert description for the new row
    bullet = f"{item.Details.ja}"
    sheet.Range[f"A{insertIndex}"].Text = bullet
    sheet.Range[f"A{insertIndex}"].Style.WrapText = True

    projects_row_offset += 1

    # Insert description for the new row
    repo = f""
    website = f"{item.Links.Website}"
    if repo != "":
        repo_link = f"{repo}\n"

        # Put new row for repo link
        insertIndex = projects_base_row + projects_row_offset + 2
        sheet.InsertRow(insertIndex)
        sheet.CopyRow(projects_template_row, sheet, insertIndex, CopyRangeOptions.All)

        # Put repo link
        urlLink = sheet.HyperLinks.Add(sheet.Range[f"A{insertIndex}"])
        urlLink.Type = HyperLinkType.Url
        urlLink.TextToDisplay = repo_link
        urlLink.Address = repo

        projects_row_offset += 1

    if website != "":
        website_link = f"ウエブサイト：{website}\n"

        # Put new row for repo link
        insertIndex = projects_base_row + projects_row_offset + 2
        sheet.InsertRow(insertIndex)
        sheet.CopyRow(projects_template_row, sheet, insertIndex, CopyRangeOptions.All)

        # Put repo link
        urlLink = sheet.HyperLinks.Add(sheet.Range[f"A{insertIndex}"])
        urlLink.Type = HyperLinkType.Url
        urlLink.TextToDisplay = website_link
        urlLink.Address = website

        projects_row_offset += 1

    projects_row_offset += 1

sheet.DeleteRow(projects_base_row + 1)

# Save cv to file
wb.SaveToFile("./out/ja.cv.xlsx", FileFormat.Version2016)
print("\033[92mSuccessfully generated ja.cv.xlsx\033[0m")

print("Setting optimal row heights, this might take a while...")

command = [
    "libreoffice",
    "--headless",
    "--nologo",
    "--nofirststartwizard",
    f"macro:///Standard.Module1.OptimalRowHeight()",
    "./out/ja.cv.xlsx",
]

# Running the command using subprocess
result = subprocess.run(command, check=True, capture_output=True, text=True)

print("\033[92mSuccessfully set optimal row heights\033[0m")


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
writer.add_page(reader.pages[1])

# Save the new PDF
with open(destination_path, "wb") as output_file:
    writer.write(output_file)

# green ansi success message for pdf file
pdf_success_message = "\033[92mSuccessfully generated ja.cv.pdf\033[0m"
print(pdf_success_message)

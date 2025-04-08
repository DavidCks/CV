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

import pprint
import shutil
import subprocess
from pypdf import PdfReader, PdfWriter
from spire.xls import *
from spire.xls.common import *
from prg.__generated__.cv_model import Item, Item1
from prg.loadCV import loadCV
from datetime import date

# read cv.json
data = loadCV("cv.json")

############
#  Resume  #
############

# load japanese resume template
wb = Workbook()
wb.LoadFromFile("./履歴書.template.xlsx")
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
work_row_offset = 0
for i, item in enumerate(data.Education):

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + work_row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[1]}{base + work_row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = item.Title.ja
    print(f"Putting {item.Title.ja} in {base_letters[2]}{base + work_row_offset}")
    work_row_offset += 2

    # Set the description
    for j, detail in enumerate(item.Details):

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
work_row_offset = 0
for i, item in enumerate(data.Work):

    # Set the date
    sheet.Range[f"{base_letters[0]}{base + work_row_offset}"].Text = item.StartYear
    print(f"Putting {item.StartYear} in {base_letters[0]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[1]}{base + work_row_offset}"].Text = item.StartMonth
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")
    sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = item.Title.ja
    print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")
    work_row_offset += 2

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
# hobbies_tech = data.Skills.Hobbies.Technology.ja

# gather other skills
skills_tech = data.Skills.Skills.Technology.ja

# combine all skills
skills = f"プログラミング: {programming_tech}\nデザイン: {design_tech}\n他のスキル: {skills_tech}\n言語: {language_tech}"

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
wb.SaveToFile("./out/履歴書.xlsx", FileFormat.Version2016)

# green ansi success message for xlsx file
xlsx_success_message = "\033[92mSuccessfully generated 履歴書.xlsx\033[0m"
print(xlsx_success_message)

# Run LibreOffice in headless mode to convert the file
subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "./out/履歴書.xlsx"], check=True
)

# Define new location
soffice_output = "履歴書.pdf"
destination_path = os.path.join("./out", os.path.basename("履歴書.pdf"))

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
pdf_success_message = "\033[92mSuccessfully generated 履歴書.pdf\033[0m"
print(pdf_success_message)

############
#    CV    #
############

print("Generating CV...")

from spire.doc import *
from spire.doc.common import *

# Create a Document object
doc = Document()

# Title style
TitleStyle = ParagraphStyle(doc)
TitleStyle.Name = "CVTitle"
TitleStyle.CharacterFormat.FontName = "ＭＳ 明朝"
TitleStyle.CharacterFormat.FontSize = 14
doc.Styles.Add(TitleStyle)


def applyTitleStyle(paragraph):
    """
    Apply the title style to the given paragraph.
    """
    paragraph.ApplyStyle(TitleStyle)
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.Format.AfterSpacing = 10
    paragraph.Format.BeforeSpacing = 10


# Text style
TextStyle = ParagraphStyle(doc)
TextStyle.Name = "CVText"
TextStyle.CharacterFormat.FontName = "ＭＳ 明朝"
TextStyle.CharacterFormat.FontSize = 10
doc.Styles.Add(TextStyle)


def applyTextStyle(paragraph):
    """
    Apply the text style to the given paragraph.
    """
    paragraph.ApplyStyle(TextStyle)
    paragraph.Format.LineSpacing = 15
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left


NameStyle = ParagraphStyle(doc)
NameStyle.Name = "CVName"
NameStyle.CharacterFormat.FontName = "ＭＳ 明朝"
NameStyle.CharacterFormat.FontSize = 10
NameStyle.CharacterFormat.UnderlineStyle = UnderlineStyle.Single
doc.Styles.Add(NameStyle)


def applyNameStyle(paragraph):
    """
    Apply the name style to the given paragraph.
    """
    paragraph.ApplyStyle(NameStyle)
    paragraph.Format.AfterSpacing = 0
    paragraph.Format.BeforeSpacing = 0


section = doc.AddSection()

section.PageSetup.Margins.All = 50
titleParagraph = section.AddParagraph()
titleParagraph.AppendText("職 務 経 歴 書")
applyTitleStyle(titleParagraph)

currentJaDate = date.today().strftime("%Y年%m月%d日")
dateParagraph = section.AddParagraph()
dateParagraph.AppendText(f"{currentJaDate}")
applyTextStyle(dateParagraph)
dateParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

name = data.Profile.Name.ja
nameParagraph = section.AddParagraph()
nameParagraph.AppendText(f"氏名　{name}")
applyNameStyle(nameParagraph)
nameParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

# Add an empty line
section.AddParagraph().AppendText("")

# Add the work summary title
workSummaryTitleParagraph = section.AddParagraph()
workSummaryTitleParagraph.AppendText("■職務要約")
applyTextStyle(workSummaryTitleParagraph)

# Add the work summary
workSummaryParagraph = section.AddParagraph()
workSummaryParagraph.AppendText(data.Skills.Programming.Details.ja)
applyTextStyle(workSummaryParagraph)

# Add an empty line
section.AddParagraph().AppendText("")

# Add skills title
skillsTitleParagraph = section.AddParagraph()
skillsTitleParagraph.AppendText("■活かせる経験・知識・技術")
applyTextStyle(skillsTitleParagraph)

# Add skills text
for i, workItem in enumerate(data.Work):
    for j, workItemDetail in enumerate(workItem.Details):
        skillsText = f"・{workItemDetail.Description.ja}"
        skillsTextParagraph = section.AddParagraph()
        skillsTextParagraph.AppendText(skillsText)
        applyTextStyle(skillsTextParagraph)

# Add an empty line
section.AddParagraph().AppendText("")

# Add work history title
workHistoryTitleParagraph = section.AddParagraph()
workHistoryTitleParagraph.AppendText("■職務経歴")
applyTextStyle(workHistoryTitleParagraph)

workHistorytable = Table(doc, True)
workHistorytable.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))
workHistorytable.TableFormat.Borders.BorderType = BorderStyle.none
workHistorytable.TableFormat.Borders.Color = Color.get_White()


def createRowImpl(
    table: Table, t1: str, t2: str, t3: str, w1: float, w2: float, w3: float
):
    row = table.AddRow(False, 3)
    row.RowFormat.Borders.BorderType = BorderStyle.none
    row.RowFormat.Borders.Color = Color.get_White()
    cell1: TableCell = row.Cells[0]
    cell2: TableCell = row.Cells[1]
    cell3: TableCell = row.Cells[2]
    cell1.SetCellWidth(w1, CellWidthType.Percentage)
    cell2.SetCellWidth(w2, CellWidthType.Percentage)
    cell3.SetCellWidth(w3, CellWidthType.Percentage)
    cell1.CellFormat.Borders.BorderType = BorderStyle.none
    cell2.CellFormat.Borders.BorderType = BorderStyle.none
    cell3.CellFormat.Borders.BorderType = BorderStyle.none
    cell1.CellFormat.Borders.Color = Color.get_White()
    cell2.CellFormat.Borders.Color = Color.get_White()
    cell3.CellFormat.Borders.Color = Color.get_White()
    cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    cell3.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph1 = cell1.AddParagraph()
    paragraph2 = cell2.AddParagraph()
    paragraph3 = cell3.AddParagraph()
    paragraph1.Format.HorizontalAlignment = HorizontalAlignment.Left
    paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Left
    paragraph3.Format.HorizontalAlignment = HorizontalAlignment.Left
    paragraph1.AppendText(t1)
    paragraph2.AppendText(t2)
    paragraph3.AppendText(t3)
    applyTextStyle(paragraph1)
    applyTextStyle(paragraph2)
    applyTextStyle(paragraph3)
    return (cell1, cell2, cell3)


def createCreateRow(table, w1=61, w2=22, w3=17):
    return lambda t1, t2, t3: createRowImpl(table, t1, t2, t3, w1, w2, w3)


createRow = createCreateRow(workHistorytable)

for i, workItem in enumerate(data.Work):
    # Generate work title
    dateJaFrom = f"{workItem.StartYear}年{workItem.StartMonth}月"
    dateJaTo = f"{workItem.EndYear}年{workItem.EndMonth}月"
    dateText = f"{dateJaFrom}～{dateJaTo}"
    titleText = f"{dateText}　　{workItem.Title.ja}"

    (cell1, cell2, cell3) = createRow(titleText, "", "")
    cell1.CellFormat.Borders.Right.BorderType = BorderStyle.none
    cell2.CellFormat.Borders.Left.BorderType = BorderStyle.none
    cell2.CellFormat.Borders.Right.BorderType = BorderStyle.none
    cell3.CellFormat.Borders.Left.BorderType = BorderStyle.none
    cell1.CellFormat.BackColor = Color.get_LightGray()
    cell2.CellFormat.BackColor = Color.get_LightGray()
    cell3.CellFormat.BackColor = Color.get_LightGray()

    # Generate company details text
    companyDescriptionText = f"事業内容：{workItem.CompanyDescription.ja}"
    companyWorkersText = f"従業員数：{workItem.Workers.ja}"
    companyDetailsText = f"{companyDescriptionText}\n{companyWorkersText}"
    roleText = f"{workItem.Role.ja}"

    (cell1, cell2, cell3) = createRow(companyDetailsText, "", roleText)
    cell1.CellFormat.Borders.Right.BorderType = BorderStyle.none
    cell2.CellFormat.Borders.Left.BorderType = BorderStyle.none
    cell2.CellFormat.Borders.Right.BorderType = BorderStyle.none

    (cell1, cell2, cell3) = createRow("", "開発環境", "規模")
    cell1.CellFormat.Borders.Bottom.BorderType = BorderStyle.none
    cell2.CellFormat.BackColor = Color.get_LightGray()
    cell3.CellFormat.BackColor = Color.get_LightGray()

    # Generate work details text
    workDetailsText = "【プロジェクト概要】\n"
    for workItemDetail in workItem.Details:
        workDetailsText += f"・{workItemDetail.Description.ja}\n"
    workDetailsText += "\n"

    # Generate work responsibilities text
    workResponsibilitiesText = "【担当フェーズ】\n"
    for workItemResponsibility in workItem.Responsibilities:
        workResponsibilitiesText += (
            f"・{workItemResponsibility.Title.ja}：{workItemResponsibility.Tasks.ja}\n"
        )
    workResponsibilitiesText += "\n"

    # Generate work achievements text
    workAchievementsText = "【実績・取り組み】\n"
    for workItemAchievement in workItem.Achievements:
        workAchievementsText += f"・{workItemAchievement.Description.ja}\n"
    workDescriptionText = (
        workDetailsText + workResponsibilitiesText + workAchievementsText
    )

    # Generate work environment text
    workEnvironmentText = ""
    for workItemEnvironment in workItem.Environment:
        workEnvironmentText += f"\n【{workItemEnvironment.Title.ja}】\n"
        for workItemEnvironmentItem in workItemEnvironment.Items:
            workEnvironmentText += f"{workItemEnvironmentItem.Item.ja}, "
        workEnvironmentText = workEnvironmentText[:-2] + "\n"
    workEnvironmentText = workEnvironmentText.rstrip()

    # Generate work scale text
    workScaleText = f"{workItem.Scope.ja}"

    (cell1, cell2, cell3) = createRow(
        workDescriptionText, workEnvironmentText, workScaleText
    )
    cell1.CellFormat.Borders.Top.BorderType = BorderStyle.none

section.Tables.Add(workHistorytable)

# Add an empty line
section.AddParagraph().AppendText("")

# Add work history title
techSkillsTitleParagraph = section.AddParagraph()
techSkillsTitleParagraph.AppendText("■テクニカルスキル")
applyTextStyle(techSkillsTitleParagraph)

techSkillsTable = Table(doc, True)
techSkillsTable.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))
techSkillsTable.TableFormat.Borders.BorderType = BorderStyle.none
techSkillsTable.TableFormat.Borders.Color = Color.get_White()

createRow = createCreateRow(techSkillsTable, 20, 45, 35)

createRow("種類", "技術", "レベル")

# categorize work evironment items by type and level
envItems: dict = {}
for i, workItem in enumerate(data.Work):
    for workItemEnvironment in workItem.Environment:
        for workItemEnvironmentItem in workItemEnvironment.Items:
            if workItemEnvironment.Title.ja not in envItems:
                envItems.update({workItemEnvironment.Title.ja: []})
            if workItemEnvironmentItem not in envItems[workItemEnvironment.Title.ja]:
                envItems[workItemEnvironment.Title.ja].append(workItemEnvironmentItem)

envItemsByLevel: dict[str, dict] = {}
for key, value in envItems.items():
    envItemsByLevel.update({key: {}})
    for envItemValue in value:
        level = envItemValue.LevelKey
        if level not in envItemsByLevel.get(key).keys():
            envItemsByLevel[key].update({level: []})
        envItemsByLevel[key][level].append(envItemValue.Item.ja)

levelsKeys = data.classes.Levels.dict().keys()

# iterate over envItemsByLevel and generate the table
for key, value in envItemsByLevel.items():
    hasTitle = False

    # Generate the title
    titleText = f"{key}"
    levelSortedItems = []

    # Sort itmes by level
    for level in levelsKeys:
        if level not in value.keys():
            continue
        levelSortedItems.append({level: value[level]})

    for item in levelSortedItems:
        for level, items in item.items():
            itemsText = ", ".join(items)

        levelText = data.classes.Levels.dict().get(level).get("ja")
        if not hasTitle:
            hasTitle = True
            (cell1, cell2, cell3) = createRow(titleText, itemsText, levelText)
            cell1.CellFormat.Borders.Bottom.BorderType = BorderStyle.none
        else:
            (cell1, cell2, cell3) = createRow("", itemsText, levelText)
            cell1.CellFormat.Borders.Top.BorderType = BorderStyle.none
            cell1.CellFormat.Borders.Bottom.BorderType = BorderStyle.none

cell1.CellFormat.Borders.Bottom.BorderType = BorderStyle.Single
section.Tables.Add(techSkillsTable)

# Add an empty line
section.AddParagraph().AppendText("")

# Add Self-PR title
selfPRTitleParagraph = section.AddParagraph()
selfPRTitleParagraph.AppendText("■自己PR")
applyTextStyle(selfPRTitleParagraph)

selfPRTextParagraph1 = section.AddParagraph()
selfPRTextParagraph1.AppendText(data.Skills.Skills.Details.ja + "\n")
applyTextStyle(selfPRTextParagraph1)

selfPRTextParagraph2 = section.AddParagraph()
selfPRTextParagraph2.AppendText("　" + data.Expectations.Workplace.ja + "\n\n")
applyTextStyle(selfPRTextParagraph2)

EndTextParagraph = section.AddParagraph()
EndTextParagraph.AppendText("\n以上")
applyTextStyle(EndTextParagraph)
EndTextParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

footer = section.HeadersFooters.Footer
footerPara = footer.AddParagraph()
footerPara.AppendField("page number", FieldType.FieldPage)
footerPara.AppendText(" / ")
footerPara.AppendField("page count", FieldType.FieldNumPages)
applyTextStyle(footerPara)
footerPara.Format.HorizontalAlignment = HorizontalAlignment.Right


doc.SaveToFile("out/職 務 経 歴 書.docx", FileFormat.Docx2019)

# green ansi success message for docx file
docx_success_message = "\033[92mSuccessfully generated 職 務 経 歴 書.docx\033[0m"

print(docx_success_message)

import subprocess


import subprocess


def run_libreoffice_macro(file_path):
    try:
        # Command to run the LibreOffice macro on the specified file
        command = [
            "soffice",
            "--invisible",
            "--nofirststartwizard",
            "--headless",
            "--norestore",
            file_path,  # Provide the file name here
            "macro:///Standard.Module1.RemoveFirstLineAndSavePDF()",
        ]

        # Run the command
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"Error executing macro: {result.stderr}")

    except Exception as e:
        print(f"Failed to run macro: {e}")


# Example file path (adjust according to your system)
file_path = "out/職 務 経 歴 書.docx"
run_libreoffice_macro(file_path)
# green ansi success message for pdf file
pdf_success_message = "\033[92mSuccessfully generated 職 務 経 歴 書.pdf\033[0m"
print(pdf_success_message)

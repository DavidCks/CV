import shutil
import subprocess
from pypdf import PdfReader, PdfWriter
from prg.__generated__.cv_model import Item, Item1
from prg.loadCV import loadCV
from datetime import date

# read cv.json
data = loadCV("cv.json")
from spire.xls import *
from spire.xls.common import *


def runResume(lang: str):

    ############
    #  Resume  #
    ############

    # load japanese resume template
    wb = Workbook()
    wb.LoadFromFile(f"./履歴書.{lang}.template.xlsx")
    sheet = wb.Worksheets[0]

    # Set the current date
    today = date.today()
    todayJa = today.strftime("%Y年%m月%d日")
    todayEn = today.strftime("%m/%d/%Y")
    todayDe = today.strftime("%d/%m/%Y")
    today = todayJa if lang == "ja" else todayEn if lang == "en" else todayDe
    sheet.Range["E3"].Text = today

    # Set the name
    sheet.Range["B5"].Text = data.Profile.Furigana
    sheet.Range["B7"].Text = data.Profile.Name.dict().get(lang)

    # Set the nationality
    sheet.Range["B13"].Text = data.Profile.Nationality.dict().get(lang)

    # Set gender
    sheet.Range["L13"].Text = data.Profile.Gender.dict().get(lang)

    # Set the birthdate
    year = data.Profile.Birthyear
    month = data.Profile.Birthmonth
    day = data.Profile.Birthday
    birthdayJa = f"{year}年{month}月{day}日"
    birthdayEn = f"{month}/{day}/{year}"
    birthdayDe = f"{day}/{month}/{year}"
    birthday = (
        birthdayJa if lang == "ja" else birthdayEn if lang == "en" else birthdayDe
    )
    sheet.Range["G13"].Text = birthday

    # Set the address
    sheet.Range["B15"].Text = data.Profile.Address.Furigana.dict().get(lang)
    sheet.Range["B17"].Text = f"〒{data.Profile.Address.Zip.dict().get(lang)}"
    sheet.Range["B19"].Text = data.Profile.Address.Address.dict().get(lang)

    # Set the phone number
    sheet.Range["K16"].Text = data.Profile.Phone.dict().get(lang)

    # Set the email
    sheet.Range["K19"].Text = data.Profile.Email

    # Fill in education title
    base = 30
    sheet.Range[f"C{base}"].Text = (
        "学歴"
        if lang == "ja"
        else "Education" if lang == "en" else "Schulisch・Akademisch"
    )

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
        sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = (
            item.Title.dict().get(lang)
        )
        print(
            f"Putting {item.Title.dict().get(lang)} in {base_letters[2]}{base + work_row_offset}"
        )
        work_row_offset += 2

        # Set the description
        for j, detail in enumerate(item.Details):

            if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 21:
                # ansi red "error" message
                error = "\033[91mYou have reached the end of the right column. Please adjust the code to handle this case.\033[0m"
                print(error)
                break

            description = f"{detail.Description.dict().get(lang)}"
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
    sheet.Range[f"{base_letters[2]}{work_title_base}"].Text = (
        "職歴"
        if lang == "ja"
        else ("Professional Experience" if lang == "en" else "Professionelle Erfahrung")
    )

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
        sheet.Range[f"{base_letters[2]}{base + work_row_offset}"].Text = (
            item.Title.dict().get(lang)
        )
        print(f"Putting {item.StartMonth} in {base_letters[1]}{base + work_row_offset}")
        work_row_offset += 2

        # Set the description
        for i, detail in enumerate(item.Details):
            if base_letters == ["O", "P", "Q"] and base + work_row_offset >= 21:
                # ansi red "error" message
                error = "\033[91mYou have reached the end of the right column in the work history. Please adjust the code to handle this case.\033[0m"
                print(error)
                break

            description = f"{detail.Description.dict().get(lang)}"
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

        description = f"{item.Title} - {item.Details.dict().get(lang)}"
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

    # gather design skills
    design_tech = data.Skills.Design.Technology.dict().get(lang)

    # gather programming skills
    programming_tech = data.Skills.Programming.Technology.dict().get(lang)

    # gather language skills
    language_tech = data.Skills.Languages.Technology.dict().get(lang)

    # gather hobbies
    # hobbies_tech = data.Skills.Hobbies.Technology.dict().get(lang)

    # gather other skills
    skills_tech = data.Skills.Skills.Technology.dict().get(lang)

    # combine all skills
    skillsJa = f"プログラミング: {programming_tech}\nデザイン: {design_tech}\n他のスキル: {skills_tech}\n言語: {language_tech}"
    skillsEn = f"Programming: {programming_tech}\nDesign: {design_tech}\nOther Skills: {skills_tech}\nLanguages: {language_tech}"
    skillsDe = f"Programmierung: {programming_tech}\nDesign: {design_tech}\nSonstige Fähigkeiten: {skills_tech}\nSprachen: {language_tech}"
    skills = skillsJa if lang == "ja" else skillsEn if lang == "en" else skillsDe

    # Fill in skills
    sheet.Range["O40"].Text = skills

    # Fill in dependents
    dependentsJa = f"{data.Profile.Dependents}人"
    dependentsEn = f"{data.Profile.Dependents} dependents"
    dependentsDe = f"{data.Profile.Dependents} Angehörige"
    dependents = (
        dependentsJa if lang == "ja" else dependentsEn if lang == "en" else dependentsDe
    )
    sheet.Range["X43"].Text = dependents
    sheet.Range["X43"].HorizontalAlignment = HorizontalAlignType.Center

    # Fill in marital status
    sheet.Range["X46"].Text = data.Profile.MaritalStatus.dict().get(lang)

    # Fill in alimony payment information
    sheet.Range["Z46"].Text = data.Profile.AlimonyPayments.dict().get(lang)

    # Fill in expectations
    workplace_expectations = data.Expectations.Workplace.dict().get(lang)
    salary_expectations = data.Expectations.Salary.dict().get(lang)
    expectationsJa = f"職場: {workplace_expectations}\n給与: {salary_expectations}"
    expectationsEn = (
        f"Workplace: {workplace_expectations}\nSalary: {salary_expectations}"
    )
    expectationsDe = (
        f"Arbeitsplatz: {workplace_expectations}\nGehalt: {salary_expectations}"
    )
    expectations = (
        expectationsJa
        if lang == "ja"
        else expectationsEn if lang == "en" else expectationsDe
    )
    sheet.Range["O51"].Text = expectations

    # save japanese cv
    wb.SaveToFile("./out/" + lang + "/履歴書.xlsx", FileFormat.Version2016)

    # green ansi success message for xlsx file
    xlsx_success_message = "\033[92mSuccessfully generated 履歴書.xlsx\033[0m"
    print(xlsx_success_message)

    # Run LibreOffice in headless mode to convert the file
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "./out/" + lang + "/履歴書.xlsx",
        ],
        check=True,
    )

    # Define new location
    soffice_output = "履歴書.pdf"
    destination_path = os.path.join("./out/" + lang, os.path.basename("履歴書.pdf"))

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

import shutil
import subprocess
from pypdf import PdfReader, PdfWriter
from prg.__generated__.cv_model import Item, Item1
from prg.loadCV import loadCV
from datetime import date

# read cv.json
data = loadCV("cv.json")

from spire.doc import *
from spire.doc.common import *


def runCV(lang: str):
    ############
    #    CV    #
    ############

    print("Generating CV...")

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
    titleTextJa = "職 務 経 歴 書"
    titleTextEn = "Curriculum Vitae"
    titleTextDe = "Résumé"
    titleText = (
        titleTextJa if lang == "ja" else titleTextEn if lang == "en" else titleTextDe
    )
    titleParagraph.AppendText(titleText)
    applyTitleStyle(titleParagraph)

    currentJaDate = date.today().strftime("%Y年%m月%d日")
    currentEnDate = date.today().strftime("%B %d, %Y")
    currentDeDate = date.today().strftime("%d. %B %Y")
    currentDate = (
        currentJaDate
        if lang == "ja"
        else currentEnDate if lang == "en" else currentDeDate
    )
    dateParagraph = section.AddParagraph()
    dateParagraph.AppendText(f"{currentDate}")
    applyTextStyle(dateParagraph)
    dateParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    name = data.Profile.Name.dict().get(lang)
    nameParagraph = section.AddParagraph()
    nameJa = f"氏名　{data.Profile.Name.dict().get('ja')}"
    nameEn = f"{data.Profile.Name.dict().get('en')}"
    nameDe = f"{data.Profile.Name.dict().get('de')}"
    name = nameJa if lang == "ja" else nameEn if lang == "en" else nameDe
    nameParagraph.AppendText(f"{name}")
    applyNameStyle(nameParagraph)
    nameParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    # Add an empty line
    section.AddParagraph().AppendText("")

    # Add the work summary title
    workSummaryTitleParagraph = section.AddParagraph()
    workSummaryTitleTextJa = "職務要約"
    workSummaryTitleTextEn = "Work Summary"
    workSummaryTitleTextDe = "Berufserfahrung"
    workSummaryTitleText = (
        workSummaryTitleTextJa
        if lang == "ja"
        else workSummaryTitleTextEn if lang == "en" else workSummaryTitleTextDe
    )
    workSummaryTitleParagraph.AppendText(f"■{workSummaryTitleText}")
    applyTextStyle(workSummaryTitleParagraph)

    # Add the work summary
    workSummaryParagraph = section.AddParagraph()
    workSummaryParagraph.AppendText(data.Skills.Programming.Details.dict().get(lang))
    applyTextStyle(workSummaryParagraph)

    # Add an empty line
    section.AddParagraph().AppendText("")

    # Add skills title
    skillsTitleParagraph = section.AddParagraph()
    skillsTitleTextJa = "■活かせる経験・知識・技術"
    skillsTitleTextEn = "■ Expertise・Experience・Skills"
    skillsTitleTextDe = "■ Expertise・Erfahrung・Fähigkeiten"
    skillsTitleText = (
        skillsTitleTextJa
        if lang == "ja"
        else skillsTitleTextEn if lang == "en" else skillsTitleTextDe
    )
    skillsTitleParagraph.AppendText(f"{skillsTitleText}")
    applyTextStyle(skillsTitleParagraph)

    # Add skills text
    for i, workItem in enumerate(data.Work):
        for j, workItemDetail in enumerate(workItem.Details):
            skillsText = f"・{workItemDetail.Description.dict().get(lang)}"
            skillsTextParagraph = section.AddParagraph()
            skillsTextParagraph.AppendText(skillsText)
            applyTextStyle(skillsTextParagraph)

    # Add an empty line
    section.AddParagraph().AppendText("")

    # Add work history title
    workHistoryTitleParagraph = section.AddParagraph()
    workHistoryTitleTextJa = "■職務経歴"
    workHistoryTitleTextEn = "■ Work History"
    workHistoryTitleTextDe = "■ Berufserfahrung"
    workHistoryTitleText = (
        workHistoryTitleTextJa
        if lang == "ja"
        else workHistoryTitleTextEn if lang == "en" else workHistoryTitleTextDe
    )

    workHistoryTitleParagraph.AppendText(f"{workHistoryTitleText}")
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
        dateEnFrom = f"{workItem.StartYear}/{workItem.StartMonth}"
        dateDeFrom = f"{workItem.StartYear}/{workItem.StartMonth}"
        dateFrom = (
            dateJaFrom if lang == "ja" else dateEnFrom if lang == "en" else dateDeFrom
        )
        dateJaTo = f"{workItem.EndYear}年{workItem.EndMonth}月"
        dateEnTo = f"{workItem.EndYear}/{workItem.EndMonth}"
        dateDeTo = f"{workItem.EndYear}/{workItem.EndMonth}"
        dateTo = dateJaTo if lang == "ja" else dateEnTo if lang == "en" else dateDeTo
        dateText = f"{dateFrom}～{dateTo}"
        titleText = f"{dateText}　　{workItem.Title.dict().get(lang)}"

        (cell1, cell2, cell3) = createRow(titleText, "", "")
        cell1.CellFormat.Borders.Right.BorderType = BorderStyle.none
        cell2.CellFormat.Borders.Left.BorderType = BorderStyle.none
        cell2.CellFormat.Borders.Right.BorderType = BorderStyle.none
        cell3.CellFormat.Borders.Left.BorderType = BorderStyle.none
        cell1.CellFormat.BackColor = Color.get_LightGray()
        cell2.CellFormat.BackColor = Color.get_LightGray()
        cell3.CellFormat.BackColor = Color.get_LightGray()

        # Generate company details text
        companyDescriptionLabelJa = "事業内容："
        companyDescriptionLabelEn = "Company Description: "
        companyDescriptionLabelDe = "Unternehmensbeschreibung: "
        companyDescriptionLabel = (
            companyDescriptionLabelJa
            if lang == "ja"
            else (
                companyDescriptionLabelEn if lang == "en" else companyDescriptionLabelDe
            )
        )
        companyDescriptionText = (
            f"{companyDescriptionLabel}{workItem.CompanyDescription.dict().get(lang)}"
        )
        companyWorkersLabelJa = "従業員数："
        companyWorkersLabelEn = "Number of Employees: "
        companyWorkersLabelDe = "Anzahl der Mitarbeiter: "
        companyWorkersLabel = (
            companyWorkersLabelJa
            if lang == "ja"
            else (companyWorkersLabelEn if lang == "en" else companyWorkersLabelDe)
        )
        companyWorkersText = f"{companyWorkersLabel}{workItem.Workers.dict().get(lang)}"
        companyDetailsText = f"{companyDescriptionText}\n{companyWorkersText}"
        roleText = f"{workItem.Role.dict().get(lang)}"

        (cell1, cell2, cell3) = createRow(companyDetailsText, "", roleText)
        cell1.CellFormat.Borders.Right.BorderType = BorderStyle.none
        cell2.CellFormat.Borders.Left.BorderType = BorderStyle.none
        cell2.CellFormat.Borders.Right.BorderType = BorderStyle.none

        envLabelJa = "開発環境"
        envLabelEn = "Development Environment"
        envLabelDe = "Entwicklungsumgebung"

        extentJa = "規模"
        extentEn = "Scale"
        extentDe = "Umfang"

        (cell1, cell2, cell3) = (
            createRow("", envLabelJa, extentJa)
            if lang == "ja"
            else (
                createRow("", envLabelEn, extentEn)
                if lang == "en"
                else createRow("", envLabelDe, extentDe)
            )
        )
        cell1.CellFormat.Borders.Bottom.BorderType = BorderStyle.none
        cell2.CellFormat.BackColor = Color.get_LightGray()
        cell3.CellFormat.BackColor = Color.get_LightGray()

        # Generate work details text
        projectLabelJa = "【プロジェクト概要】"
        projectLabelEn = " [Project Overview]"
        projectLabelDe = " [Projektübersicht]"
        projectLabel = (
            projectLabelJa
            if lang == "ja"
            else (projectLabelEn if lang == "en" else projectLabelDe)
        )
        workDetailsText = f"{projectLabel}\n"
        for workItemDetail in workItem.Details:
            workDetailsText += f"・{workItemDetail.Description.dict().get(lang)}\n"
        workDetailsText += "\n"

        # Generate work responsibilities text
        workResponsibilitiesLabelJa = "【担当業務】"
        workResponsibilitiesLabelEn = " [Responsibilities]"
        workResponsibilitiesLabelDe = " [Verantwortlichkeiten]"
        workResponsibilitiesLabel = (
            workResponsibilitiesLabelJa
            if lang == "ja"
            else (
                workResponsibilitiesLabelEn
                if lang == "en"
                else workResponsibilitiesLabelDe
            )
        )
        workResponsibilitiesText = f"{workResponsibilitiesLabel}\n"
        for workItemResponsibility in workItem.Responsibilities:
            workResponsibilitiesText += f"・{workItemResponsibility.Title.dict().get(lang)} - {workItemResponsibility.Tasks.dict().get(lang)}\n"
        workResponsibilitiesText += "\n"

        # Generate work achievements text
        workAchievementsLabelJa = "【実績・取り組み】"
        workAchievementsLabelEn = " [Achievements]"
        workAchievementsLabelDe = " [Erfolge]"
        workAchievementsLabel = (
            workAchievementsLabelJa
            if lang == "ja"
            else (workAchievementsLabelEn if lang == "en" else workAchievementsLabelDe)
        )
        workAchievementsText = f"{workAchievementsLabel}\n"
        for workItemAchievement in workItem.Achievements:
            workAchievementsText += (
                f"・{workItemAchievement.Description.dict().get(lang)}\n"
            )
        workDescriptionText = (
            workDetailsText + workResponsibilitiesText + workAchievementsText
        )

        # Generate work environment text
        workEnvironmentText = ""
        for workItemEnvironment in workItem.Environment:
            workEnvironmentText += (
                f"\n【{workItemEnvironment.Title.dict().get(lang)}】\n"
            )
            for workItemEnvironmentItem in workItemEnvironment.Items:
                workEnvironmentText += (
                    f"{workItemEnvironmentItem.Item.dict().get(lang)}, "
                )
            workEnvironmentText = workEnvironmentText[:-2] + "\n"
        workEnvironmentText = workEnvironmentText.rstrip()

        # Generate work scale text
        workScaleText = f"{workItem.Scope.dict().get(lang)}"

        (cell1, cell2, cell3) = createRow(
            workDescriptionText, workEnvironmentText, workScaleText
        )
        cell1.CellFormat.Borders.Top.BorderType = BorderStyle.none

    section.Tables.Add(workHistorytable)

    # Add an empty line
    section.AddParagraph().AppendText("")

    # Add work history title
    techSkillsTitleParagraph = section.AddParagraph()
    techSkillsTitleTextJa = "■活かせる技術"
    techSkillsTitleTextEn = "■ Technical Skills"
    techSkillsTitleTextDe = "■ Technische Fähigkeiten"
    techSkillsTitleText = (
        techSkillsTitleTextJa
        if lang == "ja"
        else techSkillsTitleTextEn if lang == "en" else techSkillsTitleTextDe
    )
    techSkillsTitleParagraph.AppendText(f"{techSkillsTitleText}")
    applyTextStyle(techSkillsTitleParagraph)

    techSkillsTable = Table(doc, True)
    techSkillsTable.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))
    techSkillsTable.TableFormat.Borders.BorderType = BorderStyle.none
    techSkillsTable.TableFormat.Borders.Color = Color.get_White()

    createRow = createCreateRow(techSkillsTable, 20, 45, 35)

    # Generate the header
    typeLabelJa = "種類"
    typeLabelEn = "Type"
    typeLabelDe = "Art"
    typeLabel = (
        typeLabelJa if lang == "ja" else (typeLabelEn if lang == "en" else typeLabelDe)
    )

    techSkillsLabelJa = "技術"
    techSkillsLabelEn = "Technical Skills"
    techSkillsLabelDe = "Technische Fähigkeiten"
    techSkillsLabel = (
        techSkillsLabelJa
        if lang == "ja"
        else (techSkillsLabelEn if lang == "en" else techSkillsLabelDe)
    )

    levelLabelJa = "レベル"
    levelLabelEn = "Level"
    levelLabelDe = "Niveau"
    levelLabel = (
        levelLabelJa
        if lang == "ja"
        else (levelLabelEn if lang == "en" else levelLabelDe)
    )
    createRow(typeLabel, techSkillsLabel, levelLabel)

    # categorize work evironment items by type and level
    envItems: dict = {}
    for i, workItem in enumerate(data.Work):
        for workItemEnvironment in workItem.Environment:
            for workItemEnvironmentItem in workItemEnvironment.Items:
                if workItemEnvironment.Title.dict().get(lang) not in envItems:
                    envItems.update({workItemEnvironment.Title.dict().get(lang): []})
                if (
                    workItemEnvironmentItem
                    not in envItems[workItemEnvironment.Title.dict().get(lang)]
                ):
                    envItems[workItemEnvironment.Title.dict().get(lang)].append(
                        workItemEnvironmentItem
                    )

    envItemsByLevel: dict[str, dict] = {}
    for key, value in envItems.items():
        envItemsByLevel.update({key: {}})
        for envItemValue in value:
            level = envItemValue.LevelKey
            if level not in envItemsByLevel.get(key).keys():
                envItemsByLevel[key].update({level: []})
            envItemsByLevel[key][level].append(envItemValue.Item.dict().get(lang))

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

            levelText = data.classes.Levels.dict().get(level).get(lang)
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
    selfPRTitleTextJa = "■自己PR"
    selfPRTitleTextEn = "■ Self-PR"
    selfPRTitleTextDe = "■ Selbst-PR"
    selfPRTitleText = (
        selfPRTitleTextJa
        if lang == "ja"
        else selfPRTitleTextEn if lang == "en" else selfPRTitleTextDe
    )
    selfPRTitleParagraph.AppendText(f"{selfPRTitleText}")
    applyTextStyle(selfPRTitleParagraph)

    selfPRTextParagraph1 = section.AddParagraph()
    selfPRTextParagraph1.AppendText(data.Skills.Skills.Details.dict().get(lang) + "\n")
    applyTextStyle(selfPRTextParagraph1)

    selfPRTextParagraph2 = section.AddParagraph()
    selfPRTextParagraph2.AppendText(
        "　"
        if lang == "ja"
        else "" + data.Expectations.Workplace.dict().get(lang) + "\n\n"
    )
    applyTextStyle(selfPRTextParagraph2)

    EndTextParagraph = section.AddParagraph()
    endTextLabelJa = "以上"
    endTextLabelEn = "End"
    endTextLabelDe = "---"
    endTextLabel = (
        endTextLabelJa
        if lang == "ja"
        else (endTextLabelEn if lang == "en" else endTextLabelDe)
    )
    EndTextParagraph.AppendText(f"\n{endTextLabel}")
    applyTextStyle(EndTextParagraph)
    EndTextParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    footer = section.HeadersFooters.Footer
    footerPara = footer.AddParagraph()
    footerPara.AppendField("page number", FieldType.FieldPage)
    footerPara.AppendText(" / ")
    footerPara.AppendField("page count", FieldType.FieldNumPages)
    applyTextStyle(footerPara)
    footerPara.Format.HorizontalAlignment = HorizontalAlignment.Right

    doc.SaveToFile("out/" + lang + "/職 務 経 歴 書.docx", FileFormat.Docx2019)

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
    file_path = "out/" + lang + "/職 務 経 歴 書.docx"
    run_libreoffice_macro(file_path)
    # green ansi success message for pdf file
    pdf_success_message = "\033[92mSuccessfully generated 職 務 経 歴 書.pdf\033[0m"
    print(pdf_success_message)

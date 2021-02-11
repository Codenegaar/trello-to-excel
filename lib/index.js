const trelloToExcell = ( () => {
    const fs = require("fs");
    const xl = require("excel4node");

    let ws, wb;
    let board;
    let lang;
    wb = new xl.Workbook();

    let startCol = 1;
    let startRow = 1;

    let listNameStyle;
    let cardNameStyle;
    let checklistHeaderStyle;
    let checklistNameStyle;
    let commentsHeaderStyle;
    let defaultStyle;

    const generateStyles = () => {
        listNameStyle = {
            alignment: { horizontal: 'center' },
            font: { bold: true },
            border: {
                top: { style: 'thick', color: "#0e0850" },
                right: { style: 'thick', color: "#0e0850" },
                bottom: { style: 'thick', color: "#0e0850" },
                left: { style: 'thick', color: "#0e0850" },
            }
        };
        cardNameStyle = {
            ...listNameStyle,
            border: {
                top: { style: 'thick', color: "#ffff00" },
                right: { style: 'thick', color: "#ffff00" },
                bottom: { style: 'thick', color: "#ffff00" },
                left: { style: 'thick', color: "#ffff00" },
            },
        };
        checklistHeaderStyle = {
            ...listNameStyle,
            border: {
                top: { style: 'thick', color: "#199897" },
                right: { style: 'thick', color: "#199897" },
                bottom: { style: 'thick', color: "#199897" },
                left: { style: 'thick', color: "#199897" },
            },
        };
        checklistNameStyle = {
            ...listNameStyle,
            border: {
                top: { style: 'thick', color: "#3a24b3" },
                right: { style: 'thick', color: "#3a24b3" },
                bottom: { style: 'thick', color: "#3a24b3" },
                left: { style: 'thick', color: "#3a24b3" },
            },
        };
        commentsHeaderStyle = {
            ...listNameStyle,
            border: {
                top: { style: 'thick', color: "#46bb4e" },
                right: { style: 'thick', color: "#46bb4e" },
                bottom: { style: 'thick', color: "#46bb4e" },
                left: { style: 'thick', color: "#46bb4e" },
            },
        }
        defaultStyle = {
            alignment: {
                horizontal: lang === 'fa' ? 'right' : 'left',
                readingOrder: lang === 'fa' ? 'rightToLeft' : 'leftToRight'
            }
        };
    };

    const readFile = path => {
        let fileData = fs.readFileSync(path);
        return JSON.parse(fileData);
    };

    const setSheetName = () => {
        ws = wb.addWorksheet(board.name);
    };

    const writeList = async (list) => {
        //Write list name
        ws.cell(startRow, startCol, startRow, startCol + 1, true).string(list.name)
            .style(listNameStyle);
        startRow += 2;

        //write cards
        for(const card of board.cards) {
            if (card.idList === list.id) {
                await writeCard(card);
            }
        }
    };

    const writeCard = async (card) => {
        //Card name
        ws.cell(startRow, startCol, startRow, startCol + 1, true).string(card.name)
            .style(cardNameStyle);
        startRow++;

        //write card info
        await writeCardAssignees(card);
        await writeDueDate(card);
        await writeCardDescription(card);
        await writeCardChecklists(card);
        await writeCardComments(card);

        //A cell is skipped
        startRow++;
    };

    const writeCardAssignees = async (card) => {
        ws.cell(startRow, startCol).string(lang === 'fa' ? "پیگیری توسط" : "Assigned to")
            .style(defaultStyle);
        let members = "";
        for (cardMemberId of card.idMembers) {
            for (boardMember of board.members) {
                if (cardMemberId === boardMember.id) {
                    members = members.concat(boardMember.username + ' - ' + boardMember.fullName + ", ");
                }
            }
        }
        ws.cell(startRow, startCol + 1).string(members)
            .style(defaultStyle);
        startRow++;
    };

    const writeDueDate = async (card) => {
        ws.cell(startRow, startCol).string(lang === 'fa' ? "ددلاین" : "Due Date")
            .style(defaultStyle);
        ws.cell(startRow, startCol + 1).string(card.due ? card.due : lang === "fa" ? "خالی" : "None")
            .style(defaultStyle);
        startRow++;
    };

    const writeCardDescription = async (card) => {
        ws.cell(startRow, startCol).string(lang === "fa" ? "توضیحات" : "Description")
            .style(defaultStyle);
        ws.cell(startRow, startCol + 1).string(card.desc)
            .style(defaultStyle);
        startRow++;
    };

    const writeCardChecklists = async (card) => {
        ws.cell(startRow, startCol, startRow, startCol + 1, true).string(lang === "fa" ? "چک لیست ها" : "Checklists")
            .style(checklistHeaderStyle);
        startRow++;

        for (const checklistId of card.idChecklists) {
            let checklist;
            for (boardChecklist of board.checklists) {
                if (boardChecklist.id === checklistId) {
                    checklist = boardChecklist;
                }
            }

            //checklist name
            ws.cell(startRow, startCol, startRow, startCol + 1, true).string(checklist.name)
                .style(checklistNameStyle);
            startRow++;

            //checklist items
            await writeChecklistItems(checklist);
        }
    };

    const writeChecklistItems = async (checklist) => {
        let doneString = lang === "fa" ? "انجام شده" : "Done";
        let notDoneString = lang === "fa" ? "جهت انجام" : "ToDo";
        for (const checkItem of checklist.checkItems) {
            ws.cell(startRow, startCol).string(checkItem.name)
                .style(defaultStyle);
                
            ws.cell(startRow, startCol + 1).string(
                checkItem.state === 'incomplete' ? notDoneString : doneString,
            ).style(defaultStyle);
            startRow++;
        }
    };

    const writeCardComments = async (card) => {
        ws.cell(startRow, startCol, startRow, startCol + 1, true).string(lang === "fa" ? 'کامنت‌ها' : "Comments")
            .style(commentsHeaderStyle);
        startRow++;

        for (const action of board.actions) {
            if (action.type === "commentCard" && action.data.card.id === card.id) {
                ws.cell(startRow, startCol).string(action.memberCreator.username + ' - ' + action.memberCreator.fullName)
                    .style(defaultStyle);

                ws.cell(startRow, startCol + 1).string(action.data.text)
                    .style(defaultStyle);

                startRow++;
            }
        } 
    };

    const convert = async (path, out, tLang) => {
        try {
            lang = tLang;
            await generateStyles();
            board = await readFile(path);
            await setSheetName();

            //Write lists
            for (const list of board.lists) {
                await writeList(list);
                startCol += 3;
                startRow = 1;
            }

            wb.write(out);
        } catch (err) {
            console.log(`Error reading file, probably doesn't exist or is invalie`);
            console.log(err);
        }
    };

    const exported = {
        convert
    };
    return exported;
} )();

module.exports = trelloToExcell;
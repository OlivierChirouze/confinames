import GoogleRange = GoogleAppsScript.Spreadsheet.Range;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

const BOARD_TAB_NAME = 'BOARD';
const WORDS_TAB_NAME = 'WORDS';
const INIT_RED = '#f4cccc';
const INIT_BLUE = '#cfe2f3';
const INIT_BLACK = '#999999';
const INIT_YELLOW = '#fcf8ec';
const NB_COLS = 5;
const NB_ROWS = 5;
const NB_CARDS_PER_COLOR = 8;
const firstCard = 'C3';
const wordFilter = 'A1';

class MathUtils {
  static getRandomInt(max: number): number {
    return Math.floor(Math.random() * Math.floor(max));
  }

  static getRandomIntButNot(max: number, except: number[]) {
    let number;
    do {
      number = MathUtils.getRandomInt(max);
    } while (except.includes(number));

    return number;
  }
}

class Ui {
  static alert(message: string) {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(message, ui.ButtonSet.OK);
  }

  static confirm(message: string): boolean {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(message, ui.ButtonSet.OK_CANCEL) === ui.Button.OK;
  }
}

class Confinames {
  doc: Spreadsheet;

  constructor() {
    this.doc = SpreadsheetApp.getActiveSpreadsheet();
  }

  private getRevealRange(card: GoogleRange) {
    return card.offset(0, 1);
  }

  private get boardTab() {
    return this.getTab(BOARD_TAB_NAME);
  }

  private get wordsTab() {
    return this.getTab(WORDS_TAB_NAME);
  }

  private getTab(tabName: string): Sheet | undefined {
    const sheets = this.doc.getSheets();
    for (const iSheet in sheets) {
      if (tabName === sheets[iSheet].getName()) {
        return sheets[iSheet];
      }
    }
    return undefined;
  }

  private getCard(sheet: Sheet, iRow: number, iCol: number) {
    return sheet.getRange(firstCard).offset(iRow * 2, iCol * 2);
  }


  private getRowCol(position: number): { row: number, col: number } {
    const row = Math.floor(position / NB_ROWS);
    const col = position % NB_ROWS;
    return {row, col};
  }

  reset() {
    const sheet = this.boardTab;

    // remove reveals
    // set words to A1, B1, etc and reset color
    for (let iCol = 0; iCol < NB_COLS; iCol++) {
      for (let iRow = 0; iRow < NB_ROWS; iRow++) {
        const letter = (iCol + 10).toString(36).toUpperCase();
        const range = this.getCard(sheet, iRow, iCol);
        range.setValue(`${letter}${iRow + 1}`);
        range.setBackground(null);
        this.getRevealRange(range).setValue(null);
      }
    }
  }

  private get wordsRange(): GoogleRange {
    const fullRange = this.wordsTab.getRange(wordFilter).getFilter().getRange();
    return fullRange.offset(1, 0, fullRange.getNumRows() - 1);
  }

  drawWords() {
    // randomly take words from the list and set it to cells
    const wordsRange = this.wordsRange;
    const sheet = this.boardTab;

    let except: number[] = [];
    const max = wordsRange.getNumRows();
    const values = wordsRange.getValues();

    const setRandomWord = (i: number) => {
      const random = MathUtils.getRandomIntButNot(max, except);
      const value = values[random][0].toString().toLocaleUpperCase();
      const position = this.getRowCol(i);
      this.getCard(sheet, position.row, position.col).setValue(value);
      except.push(random);
    }

    for (let i = 0; i < NB_ROWS * NB_COLS; i ++) {
      setRandomWord(i);
    }
  }

  drawColors() {
    const sheet = this.boardTab;

    // randomly chose which color starts
    const startColor = Math.random() >= 0.5 ? INIT_BLUE : INIT_RED;

    // randomly chose cells for red, blue, and black
    let except: number[] = [];
    const max = NB_ROWS * NB_COLS;
    const setRandomColor = (color: string) => {
      const random = MathUtils.getRandomIntButNot(max, except);
      const position = this.getRowCol(random);
      this.getCard(sheet, position.row, position.col).setBackground(color);
      except.push(random);
    }

    // First black card
    setRandomColor(INIT_BLACK);

    for (let i = 0; i < NB_CARDS_PER_COLOR; i++) {
      setRandomColor(INIT_RED);
      setRandomColor(INIT_BLUE);
    }

    // One extra card for the team that starts
    setRandomColor(startColor);

    // And now complete with yellow cards
    const remainingCards = NB_COLS * NB_ROWS - except.length;
    for (let i = 0; i < remainingCards; i ++) {
      setRandomColor(INIT_YELLOW);
    }

    Ui.alert(`Les ${startColor == INIT_RED ? 'ROUGE' : 'BLEU'} commencent !`);
  }

  revealCard(position: GoogleRange) {
    // get position current format and set next position according to format
    if (!Ui.confirm(`Révéler "${position.getValue()}" ?`))
      return;

    const nextRange = this.getRevealRange(position);
    switch (position.getBackground()) {
      case INIT_BLACK:
        nextRange.setValue('X');
        break;
      case INIT_BLUE:
        nextRange.setValue('B');
        break;
      case INIT_RED:
        nextRange.setValue('R');
        break;
      case INIT_YELLOW:
      default:
        nextRange.setValue('Y');
        break;
    }
  }
}

const mainObject = new Confinames();

function revealCurrentCard() {
  const currentRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  mainObject.revealCard(currentRange);
}

function reset() {
  mainObject.reset();
}

function newGame() {
  mainObject.reset();
  mainObject.drawWords();
  mainObject.drawColors();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('JEU')
    .addItem('Nouveau', 'newGame')
    .addItem('RAZ', 'reset')
    .addToUi();
}

/**
 * ### Description
 * Create a new rich text including no texts and no text styles.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * @param {string} text if you want to include texts, please put this.
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function createNewRichText(text = "") {
  return SpreadsheetApp.newRichTextValue().setText(text).build();
}

/**
 * ### Description
 * Append texts to existing rich text with keeping the rich text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * ### Sample input value
 * ```
 * const object = { sourceRichTextValue: SpreadsheetApp.RichTextValue, appendRichTextValue: SpreadsheetApp.RichTextValue, lastLineBreak: Boolean }
 * ```
 * - sourceRichTextValue: Source RichTextValue. This rich text is inserted to the destination rich text.
 * - appendRichTextValue: This rich text is appended to sourceRichTextValue.
 * - lastLineBreak: When this is true, when a text is appended, the line break is appended to the existing text and append the text. When you want to append a text as a paragraph, true is useful. Default value is true.
 * 
 * @param {object} object 
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function appendTexts(object) {
  if (!["sourceRichTextValue", "appendRichTextValue"].every(e => e in object && object[e].toString() == "RichTextValue")) {
    throw new Error("Invalid object.");
  }
  if (!("lastLineBreak" in object)) {
    object.lastLineBreak = true;
  }
  return new Main().appendTexts(object);
}

/**
 * ### Description
 * Insert texts as paragraphs with keeping the rich text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * ### Sample input value
 * ```
 * const object = { insertIndexAsParagraph: 0, sourceRichTextValue: SpreadsheetApp.RichTextValue, destinationRichTextValue: SpreadsheetApp.RichTextValue }
 * ```
 * - insertIndexAsParagraph: index of paragraph for inserting the rich text. The start number is 0.
 * - sourceRichTextValue: Source RichTextValue. This rich text is inserted to the destination rich text.
 * - destinationRichTextValue: Destination RichTextValue.
 * 
 * @param {object} object 
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function insertParagraphs(object) {
  if (!["sourceRichTextValue", "destinationRichTextValue"].every(e => e in object && object[e].toString() == "RichTextValue") || !("insertIndexAsParagraph" in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().insertTextsParagraphs(object);
}

/**
 * ### Description
 * Insert texts with keeping the rich text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * ### Sample input value
 * ```
 * const object = { insertIndexAsText: 0, sourceRichTextValue: SpreadsheetApp.RichTextValue, destinationRichTextValue: SpreadsheetApp.RichTextValue }
 * ```
 * - insertIndexAsText: index of text for inserting the rich text. The start number is 0.
 * - sourceRichTextValue: Source RichTextValue. This rich text is inserted to the destination rich text.
 * - destinationRichTextValue: Destination RichTextValue.
 * 
 * @param {object} object 
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function insertTexts(object) {
  if (!["sourceRichTextValue", "destinationRichTextValue"].every(e => e in object && object[e].toString() == "RichTextValue") || !("insertIndexAsText" in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().insertTextsParagraphs(object);
}

/**
 * ### Description
 * Delete paragraphs with keeping the rich text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * ### Sample input value
 * ```
 * const object = { richTextValue: SpreadsheetApp.RichTextValue, deleteIndexes: number[] }
 * ```
 * - deleteIndexes: If it's [1, 3]. The 2nd and 4th paragraphs are deleted. The start number is 0.
 * 
 * @param {object} object 
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function deleteParagraphs(object) {
  if (!("richTextValue" in object) || object.richTextValue.toString() != "RichTextValue" || !("deleteIndexes" in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().deleteParagraphs(object);
}

/**
 * ### Description
 * Delete text with keeping the rich text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * ### Sample input value
 * ```
 * const object = { richTextValue: SpreadsheetApp.RichTextValue, deleteIndexes: [{ startIndex: 1, endIndex: 3 }, , ,] }
 * ```
 * - deleteIndexes: Please set the start index and end index you want to delete in the rich text. The start number is 0.
 * 
 * @param {object} object 
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function deleteTexts(object) {
  if (!("richTextValue" in object) || object.richTextValue.toString() != "RichTextValue" || !("deleteIndexes" in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().deleteTexts(object);
}

/**
 * ### Description
 * Convert RichText object to JSON object.
 * When you want to convert RichText object to JSON object, please use convertJSONToRichText().
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * @param {SpreadsheetApp.RichTextValue} richTextValue
 * @return {object} object JSON object converted RichText object.
 */
function convertRichTextToJSON(richTextValue) {
  if (richTextValue.toString() != "RichTextValue") {
    throw new Error("Invalid object.");
  }
  return new Main().convertRichTextToJSON(richTextValue);
}

/**
 * ### Description
 * Convert JSON object to RichText object.
 * When you want to convert RichText object to JSON object, please use convertRichTextToJSON().
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 *
 * @param {object} object JSON object converted from richTextValue using convertRichTextToJSON method.
 * @return {SpreadsheetApp.RichTextValue} richTextValue
 */
function convertJSONToRichText(object) {
  if (!("obj" in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().convertJSONToRichText(object);
}

/**
 * ### Description
 * Set text style to a part of the cell text.
 * 
 * ### Sample script
 * You can see the sample script at the repository.
 * https://github.com/tanaikech/RichTextAssistant
 * 
 * ### Sample input value
 * ```
 * const object = { richText: SpreadsheetApp.RichTextValue, texts: string[], textStyle: SpreadsheetApp.TextStyle }
 * ```
 * - string[]: Regex can be used for searching texts. But, in that case, please use the string pattern instead of the regular expression literal. [Ref](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/RegExp)
 * 
 * @param {Object} object
 * @return {SpreadsheetApp.RichTextValue} SpreadsheetApp.RichTextValue
 */
function setTextStyleInCellText(object) {
  if (!["richText", "texts", "textStyle"].some(e => e in object)) {
    throw new Error("Invalid object.");
  }
  return new Main().setTextStyleInCellText(object);
}

class Main {

  appendTexts({ sourceRichTextValue, appendRichTextValue, lastLineBreak }) {
    const srcTextLen = sourceRichTextValue.getText().length;
    if (srcTextLen == 0) {
      return appendRichTextValue;
    }
    const obj1 = { insertIndexAsText: srcTextLen, sourceRichTextValue: SpreadsheetApp.newRichTextValue().setText(lastLineBreak === true ? "\n" : "").build(), destinationRichTextValue: sourceRichTextValue };
    const res = this.insertTextsParagraphs(obj1);
    const obj = { insertIndexAsText: srcTextLen + 1, sourceRichTextValue: appendRichTextValue, destinationRichTextValue: res };
    return this.insertTextsParagraphs(obj);
  }

  insertTextsParagraphs(obj) {
    const { insertIndexAsText, insertIndexAsParagraph, sourceRichTextValue, destinationRichTextValue } = obj;
    const insertText = sourceRichTextValue.getText();
    const insertTextLen = (sourceRichTextValue && insertText) ? insertText.length : 0;
    const currentText = destinationRichTextValue.getText();
    const currentTextLen = (destinationRichTextValue && currentText) ? currentText.length : 0;
    if (insertTextLen == 0) {
      return destinationRichTextValue;
    }
    if (currentTextLen == 0) {
      return sourceRichTextValue;
    }
    const insertStyles = [...insertText].map((s, i) =>
    ({
      text: s,
      link: sourceRichTextValue.getLinkUrl(i, i + 1),
      style: sourceRichTextValue.getTextStyle(i, i + 1)
    }));
    const currentStyles = [...currentText].map((s, i) =>
    ({
      text: s,
      link: destinationRichTextValue.getLinkUrl(i, i + 1),
      style: destinationRichTextValue.getTextStyle(i, i + 1)
    }));
    if ("insertIndexAsParagraph" in obj) {
      const m = [...currentText.matchAll(/\n/g)];
      const insertIndex = insertIndexAsParagraph == 0 ? -1 : m.length < insertIndexAsParagraph ? currentTextLen : m[insertIndexAsParagraph - 1].index;

      if (insertIndex == currentTextLen) {
        currentStyles.push({ text: "\n", link: null, style: SpreadsheetApp.newTextStyle().build() });
      } else {
        insertStyles.push({ text: "\n", link: null, style: SpreadsheetApp.newTextStyle().build() });
      }
      currentStyles.splice(insertIndex + 1, 0, ...insertStyles);
    } else if ("insertIndexAsText" in obj) {
      currentStyles.splice(insertIndexAsText, 0, ...insertStyles);
    }
    const text = currentStyles.map(({ text }) => text).join("");
    const newRichTextValue = SpreadsheetApp.newRichTextValue().setText(text);
    currentStyles.forEach(({ link, style }, i) =>
      newRichTextValue.setTextStyle(i, i + 1, style).setLinkUrl(i, i + 1, link)
    );
    return newRichTextValue.build();
  }

  mergeSplittedRichTextObjects_(ar) {
    const text = ar.map(({ text }) => text).join("");
    const newRichTextValue = SpreadsheetApp.newRichTextValue().setText(text);
    ar.forEach(({ link, style }, i) =>
      newRichTextValue.setTextStyle(i, i + 1, style).setLinkUrl(i, i + 1, link)
    );
    return newRichTextValue.build();
  }

  splitRichTextObjectAsEachCharacter_(richTextValue) {
    const text = richTextValue.getText();
    if (text.length == 0) return [];
    return [...text].map((s, i) => ({ text: s, link: richTextValue.getLinkUrl(i, i + 1), style: richTextValue.getTextStyle(i, i + 1) }));
  }

  deleteRichTextObject_({ richTextValue, deleteIndexes }) {
    const ar1 = this.splitRichTextObjectAsEachCharacter_(richTextValue);
    const ar2 = ar1.filter((_, i) => !deleteIndexes.some(({ startIndex, endIndex }) => i >= startIndex && i <= endIndex));
    return ar2;
  }

  deleteParagraphs({ richTextValue, deleteIndexes }) {
    const f = (n, s, i) => {
      const temp = s[i];
      if (!temp) {
        throw new Error("Paragraph for deleting is not found.");
      }
      return n += temp.length;
    };
    const text = richTextValue.getText();
    const s = text.split("\n").map((e, i, a) => i != a.length - 1 ? `${e}#` : e);
    const t = [...Array(s.length)].map((_, i) => i);
    if (([0, 1].includes(s.length) && deleteIndexes[0] == 0) || t.every(e => deleteIndexes.includes(e))) {
      return SpreadsheetApp.newRichTextValue().setText("").build();
    }
    const deleteIndexesObj = deleteIndexes.map(e => {
      const startIndex = [...Array(e)].reduce((n, _, i) => f(n, s, i), 0);
      const endIndex = [...Array(e + 1)].reduce((n, _, i) => f(n, s, i), 0) - 1;
      return { startIndex, endIndex };
    });
    if (deleteIndexes.find(e => e == s.length - 1)) {
      const ar2 = this.deleteRichTextObject_({ richTextValue, deleteIndexes: deleteIndexesObj });
      for (let i = ar2.length - 1; i >= 0; i--) {
        if (ar2[i].text.trim() == "") {
          ar2.pop();
        } else {
          break;
        }
      }
      return this.mergeSplittedRichTextObjects_(ar2);
    }
    return this.deleteTexts({ richTextValue, deleteIndexes: deleteIndexesObj });
  }

  deleteTexts({ richTextValue, deleteIndexes }) {
    const check = deleteIndexes.find(({ startIndex, endIndex }) => startIndex > endIndex);
    if (check) {
      throw new Error("endIndex is smaller than startIndex.");
    }
    const ar2 = this.deleteRichTextObject_({ richTextValue, deleteIndexes });
    return this.mergeSplittedRichTextObjects_(ar2);
  }

  convertRichTextToJSON(richTextValue) {
    const text = richTextValue.getText();
    if (text.length == 0) return { obj: [] };
    return {
      obj: [...text].map((s, i) => {
        const style = richTextValue.getTextStyle(i, i + 1);
        const params = ["getFontFamily", "getFontSize", "getForegroundColor", "isBold", "isItalic", "isStrikethrough", "isUnderline"];
        return {
          text: s,
          link: richTextValue.getLinkUrl(i, i + 1),
          style: params.reduce((o, e) => (o[e.replace(/^get|is/, "")] = style[e](), o), {})
        };
      })
    };
  }

  convertJSONToRichText(object) {
    const text = object.obj.map(({ text }) => text).join("");
    const rt = SpreadsheetApp.newRichTextValue().setText(text);
    object.obj.forEach(({ link, style }, i) => {
      const st = SpreadsheetApp.newTextStyle();
      Object.entries(style).forEach(([k, v]) => st[`set${k}`](v));
      rt.setTextStyle(i, i + 1, st.build()).setLinkUrl(i, i + 1, link);
    });
    return rt.build();
  }

  setTextStyleInCellText({ richText, texts, textStyle }) {
    const tempText = richText.getText();
    const copied = richText.copy();
    texts.forEach(t =>
      [...tempText.matchAll(new RegExp(t, "g"))].forEach(e =>
        copied.setTextStyle(e.index, e.index + e[0].length, textStyle)
      ));
    return copied.build();
  }
}

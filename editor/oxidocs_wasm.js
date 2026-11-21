/* @ts-self-types="./oxidocs_wasm.d.ts" */

/**
 * A document held in the model for editing by position.
 *
 * Positions are `{story?, para, offset}`: a paragraph numbered in document
 * order (through tables) within a story (`"body"` by default, or
 * `{"header_footer": i}`, `{"footnote": i}`, `{"endnote": i}`,
 * `{"comment": i}`), and an offset counted in Word character positions.
 * `paragraphs()` lists each paragraph's characters in that count.
 */
export class DocxDocument {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        DocxDocumentFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_docxdocument_free(ptr, 0);
    }
    /**
     * Apply edits given as a JSON array (`oxidocs_model::edit::Edit`:
     * typing, deleting, splitting and joining paragraphs, formatting, and
     * the structural edits that make any document from an empty one). All
     * apply or, on an error, none.
     * @param {string} edits_json
     */
    apply(edits_json) {
        const ptr0 = passStringToWasm0(edits_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docxdocument_apply(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * The edits Word's New Comment makes over a model range (`{para,
     * offset}` each; equal for the word at a caret), by `author` at `date`
     * (ISO 8601), as JSON `{"edits": [...], "end": {para, offset}}`.
     * @param {string} from_json
     * @param {string} to_json
     * @param {string} text
     * @param {string} author
     * @param {string} initials
     * @param {string} date
     * @returns {string}
     */
    comment_edits(from_json, to_json, text, author, initials, date) {
        let deferred8_0;
        let deferred8_1;
        try {
            const ptr0 = passStringToWasm0(from_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(to_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ptr3 = passStringToWasm0(author, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len3 = WASM_VECTOR_LEN;
            const ptr4 = passStringToWasm0(initials, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len4 = WASM_VECTOR_LEN;
            const ptr5 = passStringToWasm0(date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len5 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_comment_edits(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2, ptr3, len3, ptr4, len4, ptr5, len5);
            var ptr7 = ret[0];
            var len7 = ret[1];
            if (ret[3]) {
                ptr7 = 0; len7 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred8_0 = ptr7;
            deferred8_1 = len7;
            return getStringFromWasm0(ptr7, len7);
        } finally {
            wasm.__wbindgen_free(deferred8_0, deferred8_1, 1);
        }
    }
    /**
     * The edits Word's Insert Field makes for `code` at a model position
     * (`{para, offset}`), showing `result`, as JSON
     * `{"edits": [...], "end": {para, offset}}`.
     * @param {string} at_json
     * @param {string} code
     * @param {string} result
     * @returns {string}
     */
    field_edits(at_json, code, result) {
        let deferred5_0;
        let deferred5_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(code, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(result, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_field_edits(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2);
            var ptr4 = ret[0];
            var len4 = ret[1];
            if (ret[3]) {
                ptr4 = 0; len4 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred5_0 = ptr4;
            deferred5_1 = len4;
            return getStringFromWasm0(ptr4, len4);
        } finally {
            wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
        }
    }
    /**
     * The edits Word's Insert Footnote makes at a model position
     * (`{para, offset}`), the note holding `text`, as JSON
     * `{"edits": [...], "end": {para, offset}}`, `end` being just after the
     * reference mark once they are applied.
     * @param {string} at_json
     * @param {string} text
     * @returns {string}
     */
    footnote_edits(at_json, text) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_footnote_edits(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * The edits that give the section holding body paragraph `para` a
     * header (or, with `footer`, a footer) holding `text`, as a JSON array.
     * @param {number} para
     * @param {boolean} footer
     * @param {string} text
     * @returns {string}
     */
    header_footer_edits(para, footer, text) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_header_footer_edits(this.__wbg_ptr, para, footer, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * The edits Word's Insert Hyperlink makes for `url` over a model range
     * (`{para, offset}` each; equal for no selection), as JSON
     * `{"edits": [...], "end": {para, offset}}`, `end` being just after the
     * link once they are applied.
     * @param {string} from_json
     * @param {string} to_json
     * @param {string} url
     * @returns {string}
     */
    link_edits(from_json, to_json, url) {
        let deferred5_0;
        let deferred5_1;
        try {
            const ptr0 = passStringToWasm0(from_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(to_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(url, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_link_edits(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2);
            var ptr4 = ret[0];
            var len4 = ret[1];
            if (ret[3]) {
                ptr4 = 0; len4 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred5_0 = ptr4;
            deferred5_1 = len4;
            return getStringFromWasm0(ptr4, len4);
        } finally {
            wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
        }
    }
    /**
     * The model position (`{para, offset}`) of a place on a laid-out page,
     * given as `{section, block, cells?, run, char_offset}`: a layout
     * element's `source_container_index`, `paragraph_index`, cell path, `run_index`
     * and `char_offset` (plus the character within the element). Null when the
     * paragraph has no counterpart in the model.
     * @param {string} core_pos_json
     * @returns {string | undefined}
     */
    locate(core_pos_json) {
        const ptr0 = passStringToWasm0(core_pos_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docxdocument_locate(this.__wbg_ptr, ptr0, len0);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        let v2;
        if (ret[0] !== 0) {
            v2 = getStringFromWasm0(ret[0], ret[1]).slice();
            wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        }
        return v2;
    }
    /**
     * The whole model as JSON (the shape `docx_to_model` returns).
     * @returns {string}
     */
    model_json() {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.docxdocument_model_json(this.__wbg_ptr);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Read a .docx. A feature the model cannot express is an error naming it.
     * @param {Uint8Array} data
     */
    constructor(data) {
        const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docxdocument_new(ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        this.__wbg_ptr = ret[0] >>> 0;
        DocxDocumentFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * The edits Word's Layout › Size and Layout › Margins make for the
     * section holding body paragraph `para`, as a JSON array. `setup_json`
     * is `{"size": [width, height]?, "margins": [top, right, bottom, left]?}`
     * in twips; the size is given for portrait and turned for a landscape
     * section.
     * @param {number} para
     * @param {string} setup_json
     * @returns {string}
     */
    page_setup_edits(para, setup_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(setup_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_page_setup_edits(this.__wbg_ptr, para, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * @param {number} para
     * @returns {string}
     */
    paragraph_text(para) {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.docxdocument_paragraph_text(this.__wbg_ptr, para);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * A story's paragraphs as a JSON array of strings, one character per
     * position. `story_json` is omitted for the body.
     * @param {string | null} [story_json]
     * @returns {string}
     */
    paragraphs(story_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            var ptr0 = isLikeNone(story_json) ? 0 : passStringToWasm0(story_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len0 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_paragraphs(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Where on the laid-out document a model position falls, as
     * `{section, block, cells, run, char_offset}`; null when unknown.
     * @param {string} model_pos_json
     * @returns {string | undefined}
     */
    place(model_pos_json) {
        const ptr0 = passStringToWasm0(model_pos_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docxdocument_place(this.__wbg_ptr, ptr0, len0);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        let v2;
        if (ret[0] !== 0) {
            v2 = getStringFromWasm0(ret[0], ret[1]).slice();
            wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        }
        return v2;
    }
    /**
     * Write the document as .docx bytes.
     * @returns {Uint8Array}
     */
    save() {
        const ret = wasm.docxdocument_save(this.__wbg_ptr);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v1;
    }
    /**
     * One body paragraph's characters, one per position (see `paragraphs`).
     * The edits Word's Insert Table makes for a table of `rows` by `cols`
     * at a model position (`{para, offset}`), as JSON
     * `{"edits": [...], "cell": para}`, `cell` being the first cell's
     * paragraph once they are applied.
     * @param {string} at_json
     * @param {number} rows
     * @param {number} cols
     * @returns {string}
     */
    table_edits(at_json, rows, cols) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_table_edits(this.__wbg_ptr, ptr0, len0, rows, cols);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * The edits Word's Backspace makes at the start of a numbered paragraph
     * at a model position (`{para, offset}`): the number comes off and the
     * text keeps the list's indent. A JSON array, empty when the paragraph
     * has no number.
     * @param {string} at_json
     * @returns {string}
     */
    unnumber_edits(at_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.docxdocument_unnumber_edits(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
}
if (Symbol.dispose) DocxDocument.prototype[Symbol.dispose] = DocxDocument.prototype.free;

/**
 * Where the engine breaks one paragraph's lines, for a box `width_pt` wide.
 *
 * The browser has been wrapping slide text with its own text layout, so what
 * a person edits on screen is not what the renderer -- or PowerPoint --
 * produces. This is the engine's own break test
 * (`pptx-master-unit-break-law`: each glyph's advance rounded to 1/8 pt, then
 * summed), running where there is no font system.
 *
 * Returns `null` when the metrics tables cannot measure the text, which is
 * most families: the tables carry the faces that were measured from their
 * real files, and a guess here would break lines PowerPoint keeps. A caller
 * that gets null must say the wrap is the browser's, not the engine's.
 *
 * `runs` is the paragraph's runs, so a line carrying a bold word is measured
 * per run rather than wholly in one weight.
 * @param {string} text
 * @param {number} font_size
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {number} width_pt
 * @param {any} runs
 * @param {any} advances
 * @returns {any}
 */
export function break_slide_paragraph(text, font_size, family, bold, italic, width_pt, runs, advances) {
    const ptr0 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.break_slide_paragraph(ptr0, len0, font_size, ptr1, len1, bold, italic, width_pt, runs, advances);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * Build a .docx from a content structure.
 * `content`: JS array of block objects:
 *   { type: "paragraph", runs: [{text, bold?, italic?, underline?, strikethrough?, font_family?, font_size?, color?}], alignment?, heading_level?, line_height? }
 *   { type: "table", rows: [[{text, bold?}]] }
 * @param {any} content
 * @returns {Uint8Array}
 */
export function build_docx(content) {
    const ret = wasm.build_docx(content);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v1;
}

/**
 * Build a .docx from content, using a template docx for styles/theme/numbering.
 * Preserves original formatting while replacing document content.
 * @param {any} content
 * @param {Uint8Array} template
 * @returns {Uint8Array}
 */
export function build_docx_with_template(content, template) {
    const ptr0 = passArray8ToWasm0(template, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.build_docx_with_template(content, ptr0, len0);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

export function clear_font_programs() {
    wasm.clear_font_programs();
}

/**
 * Create a blank .docx file and return it as bytes.
 * Can be used to create a new document from scratch.
 * @returns {Uint8Array}
 */
export function create_blank_docx() {
    const ret = wasm.create_blank_docx();
    var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v1;
}

/**
 * Apply structural edits to a .docx file.
 *
 * `data`: original .docx bytes
 * `edits`: JS array of edit operation objects. Each object has a `type` field:
 *
 * Text operations:
 *   { type: "set_run_text", paragraph_index, run_index, new_text }
 *   { type: "insert_paragraph", index, text, style?, para_style? }
 *   { type: "delete_paragraph", index }
 *   { type: "insert_run", paragraph_index, run_index, text, style? }
 *   { type: "delete_run", paragraph_index, run_index }
 *
 * Formatting:
 *   { type: "set_run_format", paragraph_index, run_index, style }
 *   { type: "set_paragraph_format", paragraph_index, style }
 *
 * Tables:
 *   { type: "insert_table", index, rows, cols, content?, col_widths_pt? }
 *   { type: "insert_table_row", table_index, row_index, cells }
 *   { type: "delete_table_row", table_index, row_index }
 *   { type: "set_cell_text", table_index, row, col, text }
 *   { type: "set_table_run_text", table_index, row, col, paragraph_index, run_index, new_text }
 *     (one run of a cell, the rest of the cell left as it stands)
 *
 * Images:
 *   { type: "insert_image", index, data (base64), width_pt, height_pt, content_type }
 *
 * style (RunProps): { bold?, italic?, underline?, font_family?, font_size?, color?, highlight? }
 * para_style (ParaProps): { alignment?, space_before?, space_after?, line_spacing?, indent_left?, style_id? }
 * The runs and table cells `edit_docx_advanced` can address, as the editor
 * counts them, as JSON: `{"paragraphs": [[run text, …], …], "tables":
 * [[[{"text", "runs", "paragraphs": [[run text, …], …]}, …cells], …rows], …]}`.
 *
 * `paragraph_index` in `set_run_text` counts the body's own top-level
 * paragraphs, which is not how the parsed document (paginated, tables folded
 * in) lays them out; a page that edits text matches its runs against this
 * before it lets them be edited, so an edit lands where it was typed.
 * @param {Uint8Array} data
 * @returns {string}
 */
export function docx_addressable(data) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docx_addressable(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * The placeholders of a .docx template with their positions on the laid-out
 * pages, as JSON: `{"pages": [[width, height], …], "fields": [{"name",
 * "page", "text": [{x, y, width, height}…], "cell": {…} | null}]}`, in points
 * from the top-left corner of the page.
 * @param {Uint8Array} data
 * @returns {string}
 */
export function docx_field_boxes(data) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docx_field_boxes(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * Build a .docx from a document model given as JSON (the shape
 * `docx_to_model` returns). With `template`, any document-wide part the
 * model leaves out (styles, numbering, settings, theme, fonts) is taken
 * from the template.
 * @param {string} model_json
 * @param {Uint8Array | null} [template]
 * @returns {Uint8Array}
 */
export function docx_from_model(model_json, template) {
    const ptr0 = passStringToWasm0(model_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    var ptr1 = isLikeNone(template) ? 0 : passArray8ToWasm0(template, wasm.__wbindgen_malloc);
    var len1 = WASM_VECTOR_LEN;
    const ret = wasm.docx_from_model(ptr0, len0, ptr1, len1);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v3 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v3;
}

/**
 * Read a .docx into the document model and return it as JSON.
 *
 * The model says everything the file says: body, sections, headers and
 * footers, notes, comments, styles, numbering, settings, theme. A feature
 * the model cannot express is an error naming it, never silently dropped.
 * @param {Uint8Array} data
 * @returns {string}
 */
export function docx_to_model(data) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.docx_to_model(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * Convert a .docx file to PDF bytes.
 * Parses the docx, runs layout, and converts positioned elements to PDF.
 * @param {Uint8Array} data
 * @returns {Uint8Array}
 */
export function docx_to_pdf(data) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.docx_to_pdf(ptr0, len0);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * Convert a .docx file to PDF bytes with `title` as the PDF's title, which
 * viewers show in place of the file name or a blob URL. An empty title
 * falls back to what `docx_to_pdf` uses.
 * @param {Uint8Array} data
 * @param {string} title
 * @returns {Uint8Array}
 */
export function docx_to_pdf_with_title(data, title) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passStringToWasm0(title, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.docx_to_pdf_with_title(ptr0, len0, ptr1, len1);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v3 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v3;
}

/**
 * Edit a .docx file and return the modified bytes.
 *
 * `data`: original .docx bytes
 * `edits`: JS array of `{paragraph_index, run_index, new_text}` objects
 *
 * Returns the modified .docx as `Uint8Array`.
 * @param {Uint8Array} data
 * @param {any} edits
 * @returns {Uint8Array}
 */
export function edit_docx(data, edits) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.edit_docx(ptr0, len0, edits);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * @param {Uint8Array} data
 * @param {any} edits
 * @returns {Uint8Array}
 */
export function edit_docx_advanced(data, edits) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.edit_docx_advanced(ptr0, len0, edits);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * Edit a .pptx file and return the modified bytes.
 * @param {Uint8Array} data
 * @param {any} edits
 * @returns {Uint8Array}
 */
export function edit_pptx(data, edits) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.edit_pptx(ptr0, len0, edits);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * Edit a .pptx and break paragraphs in it, returning the modified bytes.
 *
 * `edits` replaces run text; `splits` cuts a paragraph in two at a character
 * offset, which is what Enter means. Both are applied to the same save, and
 * the text edit lands first so a split counts the characters the file will
 * actually carry. `geoms` moves and sizes top-level shapes.
 * @param {Uint8Array} data
 * @param {any} edits
 * @param {any} splits
 * @param {any} merges
 * @param {any} formats
 * @param {any} geoms
 * @param {any} aligns
 * @returns {Uint8Array}
 */
export function edit_pptx_with_splits(data, edits, splits, merges, formats, geoms, aligns) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.edit_pptx_with_splits(ptr0, len0, edits, splits, merges, formats, geoms, aligns);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * Fast text edit + re-layout using cached document (skips docx parse).
 * Returns layout result. Also updates the cached docx bytes.
 * @param {number} paragraph_index
 * @param {number} run_index
 * @param {string} new_text
 * @returns {any}
 */
export function edit_text_and_relayout(paragraph_index, run_index, new_text) {
    const ptr0 = passStringToWasm0(new_text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.edit_text_and_relayout(paragraph_index, run_index, ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * The catalog members requested since the last call that were not
 * registered: `[{ key, offset, length }]`, byte ranges in
 * `font_catalog_metrics.gz`. Empty for a bundled build.
 * @returns {any}
 */
export function font_catalog_missing() {
    const ret = wasm.font_catalog_missing();
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * Register one member's bytes (exactly its range in the catalog file).
 * Returns false when the key is unknown or the length does not match.
 * @param {string} key
 * @param {Uint8Array} bytes
 * @returns {boolean}
 */
export function font_catalog_register(key, bytes) {
    const ptr0 = passStringToWasm0(key, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.font_catalog_register(ptr0, len0, ptr1, len1);
    return ret !== 0;
}

/**
 * A successful registration changes both layout and painting together.
 * Callers use this revision to relayout after asynchronous font loading.
 * @returns {bigint}
 */
export function font_program_revision() {
    const ret = wasm.font_program_revision();
    return BigInt.asUintN(64, ret);
}

/**
 * Painting commands are returned to the renderer in memory, independently
 * from layout/source dumps. They are never stored as font data tables.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {number} index
 * @param {Float32Array} expected_bounds_em
 * @returns {any}
 */
export function get_font_glyph_outline(family, bold, italic, index, expected_bounds_em) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passArrayF32ToWasm0(expected_bounds_em, wasm.__wbindgen_malloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.get_font_glyph_outline(ptr0, len0, bold, italic, index, ptr1, len1);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * Complete SFNT bytes for the CSS face; unlike PDF embedding, preserve the
 * OTF wrapper around CFF tables. Returned bytes stay in the client memory.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @returns {Uint8Array}
 */
export function get_registered_font_sfnt(family, bold, italic) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.get_registered_font_sfnt(ptr0, len0, bold, italic);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @returns {boolean}
 */
export function has_font_program(family, bold, italic) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.has_font_program(ptr0, len0, bold, italic);
    return ret !== 0;
}

export function init() {
    wasm.init();
}

/**
 * Load a document, cache it, and return layout result.
 * Subsequent calls to `edit_text_and_relayout` will reuse the cached parse.
 * @param {Uint8Array} data
 * @returns {any}
 */
export function layout_document(data) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.layout_document(ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * Lay out one text shape the way the engine does, for the browser to draw.
 *
 * `shape` is a `Shape` from `parse_presentation`, `paragraphs` its
 * paragraphs, and `master` / `ph_levels` the inherited outline levels. The
 * answer is one entry per LINE -- its text, where it starts, its baseline,
 * and which paragraph and character offset it came from, so a click can be
 * mapped back to a run.
 *
 * `complete` says whether EVERY paragraph was measured. A shape that comes
 * back incomplete must not be drawn as if it were the engine's layout: the
 * measured tables cover 17 of the 142 families the corpora name, and the rest
 * are still the browser's own wrap.
 * @param {any} shape
 * @param {any} paragraphs
 * @param {any} master
 * @param {any} ph_levels
 * @param {string} default_family
 * @param {any} advances
 * @returns {any}
 */
export function layout_slide_shape(shape, paragraphs, master, ph_levels, default_family, advances) {
    const ptr0 = passStringToWasm0(default_family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.layout_slide_shape(shape, paragraphs, master, ph_levels, ptr0, len0, advances);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * @param {Uint8Array} data
 * @returns {any}
 */
export function parse_document(data) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.parse_document(ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * @param {Uint8Array} data
 * @returns {any}
 */
export function parse_presentation(data) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.parse_presentation(ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * Register caller-supplied font bytes in memory. The face's own name table
 * and style metadata must agree; no OS paths or family exception lists.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {Uint8Array} bytes
 * @param {number} face_index
 */
export function register_font_program(family, bold, italic, bytes, face_index) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.register_font_program(ptr0, len0, bold, italic, ptr1, len1, face_index);
    if (ret[1]) {
        throw takeFromExternrefTable0(ret[0]);
    }
}

/**
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {Uint8Array} bytes
 */
export function register_font_program_family(family, bold, italic, bytes) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.register_font_program_family(ptr0, len0, bold, italic, ptr1, len1);
    if (ret[1]) {
        throw takeFromExternrefTable0(ret[0]);
    }
}

/**
 * Write comments into a .docx (adds only). See `update_docx_comments` for
 * the full operation set (add + remove + resolve).
 * @param {Uint8Array} data
 * @param {any} comments
 * @returns {Uint8Array}
 */
export function set_docx_comments(data, comments) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.set_docx_comments(ptr0, len0, comments);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

/**
 * What the compiled tables say one character advances, in EM units.
 *
 * The page measures faces itself; this lets it check that measuring against a
 * face the tables also carry, instead of trusting a canvas it never verified.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {string} ch
 * @returns {number | undefined}
 */
export function slide_face_advance(family, bold, italic, ch) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passStringToWasm0(ch, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.slide_face_advance(ptr0, len0, bold, italic, ptr1, len1);
    return ret === 0x100000001 ? undefined : ret;
}

/**
 * Whether the engine can measure this family at all, so a caller can tell a
 * person which text on the page is laid out by the engine and which is not.
 * @param {string} family
 * @returns {boolean}
 */
export function slide_family_measurable(family) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.slide_family_measurable(ptr0, len0);
    return ret !== 0;
}

/**
 * Where each character of one laid-out line starts, in points from its `x`.
 *
 * `line` is a `PlacedLine` from [`layout_slide_shape`] and `advances` the same
 * measured faces that call was given. The answer is the engine's own placement
 * -- each advance on the master unit, which is what PowerPoint measures and
 * draws on (see `glyph_offsets_pt`) -- so a page that draws from it puts the
 * glyphs where the break already assumed they were.
 *
 * Null when a face cannot be measured; the caller then has to fall back to its
 * own measuring and should say the answer is not the engine's.
 * @param {any} line
 * @param {any} advances
 * @returns {any}
 */
export function slide_glyph_offsets(line, advances) {
    const ret = wasm.slide_glyph_offsets(line, advances);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * The face Oxi ships to draw `family` when the machine does not have it:
 * Carlito for Calibri, Liberation for Arial / Times New Roman / Courier New,
 * Caladea for Cambria, OxiMincho / OxiGothic for a Japanese family. Null
 * for a family with no stand-in.
 *
 * It is for painting only. Registering it as `family`'s program would lay
 * the text out with the stand-in's widths instead of the measured ones, so a
 * caller hands it to CSS (`FontFace`) and the layout keeps its own metrics.
 * The Japanese pair has no bold or italic member; the browser slants or
 * emboldens the regular one.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @returns {Uint8Array | undefined}
 */
export function stand_in_font_program(family, bold, italic) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.stand_in_font_program(ptr0, len0, bold, italic);
    let v2;
    if (ret[0] !== 0) {
        v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    }
    return v2;
}

/**
 * Resolve the requested name/style from the actual program's name tables,
 * including every collection member. No filename or family exception map.
 * @param {string} family
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {Uint8Array} bytes
 * @returns {boolean}
 */
export function try_register_font_program_family(family, bold, italic, bytes) {
    const ptr0 = passStringToWasm0(family, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ptr1 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
    const len1 = WASM_VECTOR_LEN;
    const ret = wasm.try_register_font_program_family(ptr0, len0, bold, italic, ptr1, len1);
    return ret !== 0;
}

/**
 * Apply a batch of comment operations to a .docx:
 * { add: [ { author, initials?, date?, text, paragraph_index, char_start,
 *            char_end, resolved?, parent_index?, parent_para_id? } ],
 *   remove_ids: [ "w:id", … ],
 *   set_resolved: [ { para_id, done } ] }
 * Adds write word/comments.xml + commentsExtended.xml (threads via
 * paraIdParent, resolved via w15:done) and range markers in document.xml;
 * removals strip all three.
 * @param {Uint8Array} data
 * @param {any} ops
 * @returns {Uint8Array}
 */
export function update_docx_comments(data, ops) {
    const ptr0 = passArray8ToWasm0(data, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.update_docx_comments(ptr0, len0, ops);
    if (ret[3]) {
        throw takeFromExternrefTable0(ret[2]);
    }
    var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
    wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
    return v2;
}

function __wbg_get_imports() {
    const import0 = {
        __proto__: null,
        __wbg_Error_83742b46f01ce22d: function(arg0, arg1) {
            const ret = Error(getStringFromWasm0(arg0, arg1));
            return ret;
        },
        __wbg_Number_a5a435bd7bbec835: function(arg0) {
            const ret = Number(arg0);
            return ret;
        },
        __wbg_String_8564e559799eccda: function(arg0, arg1) {
            const ret = String(arg1);
            const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
        },
        __wbg___wbindgen_bigint_get_as_i64_447a76b5c6ef7bda: function(arg0, arg1) {
            const v = arg1;
            const ret = typeof(v) === 'bigint' ? v : undefined;
            getDataViewMemory0().setBigInt64(arg0 + 8 * 1, isLikeNone(ret) ? BigInt(0) : ret, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, !isLikeNone(ret), true);
        },
        __wbg___wbindgen_boolean_get_c0f3f60bac5a78d1: function(arg0) {
            const v = arg0;
            const ret = typeof(v) === 'boolean' ? v : undefined;
            return isLikeNone(ret) ? 0xFFFFFF : ret ? 1 : 0;
        },
        __wbg___wbindgen_debug_string_5398f5bb970e0daa: function(arg0, arg1) {
            const ret = debugString(arg1);
            const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
        },
        __wbg___wbindgen_in_41dbb8413020e076: function(arg0, arg1) {
            const ret = arg0 in arg1;
            return ret;
        },
        __wbg___wbindgen_is_bigint_e2141d4f045b7eda: function(arg0) {
            const ret = typeof(arg0) === 'bigint';
            return ret;
        },
        __wbg___wbindgen_is_function_3c846841762788c1: function(arg0) {
            const ret = typeof(arg0) === 'function';
            return ret;
        },
        __wbg___wbindgen_is_null_0b605fc6b167c56f: function(arg0) {
            const ret = arg0 === null;
            return ret;
        },
        __wbg___wbindgen_is_object_781bc9f159099513: function(arg0) {
            const val = arg0;
            const ret = typeof(val) === 'object' && val !== null;
            return ret;
        },
        __wbg___wbindgen_is_string_7ef6b97b02428fae: function(arg0) {
            const ret = typeof(arg0) === 'string';
            return ret;
        },
        __wbg___wbindgen_is_undefined_52709e72fb9f179c: function(arg0) {
            const ret = arg0 === undefined;
            return ret;
        },
        __wbg___wbindgen_jsval_eq_ee31bfad3e536463: function(arg0, arg1) {
            const ret = arg0 === arg1;
            return ret;
        },
        __wbg___wbindgen_jsval_loose_eq_5bcc3bed3c69e72b: function(arg0, arg1) {
            const ret = arg0 == arg1;
            return ret;
        },
        __wbg___wbindgen_number_get_34bb9d9dcfa21373: function(arg0, arg1) {
            const obj = arg1;
            const ret = typeof(obj) === 'number' ? obj : undefined;
            getDataViewMemory0().setFloat64(arg0 + 8 * 1, isLikeNone(ret) ? 0 : ret, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, !isLikeNone(ret), true);
        },
        __wbg___wbindgen_string_get_395e606bd0ee4427: function(arg0, arg1) {
            const obj = arg1;
            const ret = typeof(obj) === 'string' ? obj : undefined;
            var ptr1 = isLikeNone(ret) ? 0 : passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len1 = WASM_VECTOR_LEN;
            getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
        },
        __wbg___wbindgen_throw_6ddd609b62940d55: function(arg0, arg1) {
            throw new Error(getStringFromWasm0(arg0, arg1));
        },
        __wbg_call_e133b57c9155d22c: function() { return handleError(function (arg0, arg1) {
            const ret = arg0.call(arg1);
            return ret;
        }, arguments); },
        __wbg_codePointAt_3925396be3b2c733: function(arg0, arg1) {
            const ret = arg0.codePointAt(arg1 >>> 0);
            return ret;
        },
        __wbg_done_08ce71ee07e3bd17: function(arg0) {
            const ret = arg0.done;
            return ret;
        },
        __wbg_entries_e8a20ff8c9757101: function(arg0) {
            const ret = Object.entries(arg0);
            return ret;
        },
        __wbg_error_a6fa202b58aa1cd3: function(arg0, arg1) {
            let deferred0_0;
            let deferred0_1;
            try {
                deferred0_0 = arg0;
                deferred0_1 = arg1;
                console.error(getStringFromWasm0(arg0, arg1));
            } finally {
                wasm.__wbindgen_free(deferred0_0, deferred0_1, 1);
            }
        },
        __wbg_fromCodePoint_4a02c6dcced1d233: function() { return handleError(function (arg0) {
            const ret = String.fromCodePoint(arg0 >>> 0);
            return ret;
        }, arguments); },
        __wbg_get_326e41e095fb2575: function() { return handleError(function (arg0, arg1) {
            const ret = Reflect.get(arg0, arg1);
            return ret;
        }, arguments); },
        __wbg_get_a8ee5c45dabc1b3b: function(arg0, arg1) {
            const ret = arg0[arg1 >>> 0];
            return ret;
        },
        __wbg_get_unchecked_329cfe50afab7352: function(arg0, arg1) {
            const ret = arg0[arg1 >>> 0];
            return ret;
        },
        __wbg_get_with_ref_key_6412cf3094599694: function(arg0, arg1) {
            const ret = arg0[arg1];
            return ret;
        },
        __wbg_instanceof_ArrayBuffer_101e2bf31071a9f6: function(arg0) {
            let result;
            try {
                result = arg0 instanceof ArrayBuffer;
            } catch (_) {
                result = false;
            }
            const ret = result;
            return ret;
        },
        __wbg_instanceof_Uint8Array_740438561a5b956d: function(arg0) {
            let result;
            try {
                result = arg0 instanceof Uint8Array;
            } catch (_) {
                result = false;
            }
            const ret = result;
            return ret;
        },
        __wbg_isArray_33b91feb269ff46e: function(arg0) {
            const ret = Array.isArray(arg0);
            return ret;
        },
        __wbg_isSafeInteger_ecd6a7f9c3e053cd: function(arg0) {
            const ret = Number.isSafeInteger(arg0);
            return ret;
        },
        __wbg_iterator_d8f549ec8fb061b1: function() {
            const ret = Symbol.iterator;
            return ret;
        },
        __wbg_length_15d3fc853a68bbbc: function(arg0) {
            const ret = arg0.length;
            return ret;
        },
        __wbg_length_b3416cf66a5452c8: function(arg0) {
            const ret = arg0.length;
            return ret;
        },
        __wbg_length_ea16607d7b61445b: function(arg0) {
            const ret = arg0.length;
            return ret;
        },
        __wbg_new_227d7c05414eb861: function() {
            const ret = new Error();
            return ret;
        },
        __wbg_new_49d5571bd3f0c4d4: function() {
            const ret = new Map();
            return ret;
        },
        __wbg_new_5f486cdf45a04d78: function(arg0) {
            const ret = new Uint8Array(arg0);
            return ret;
        },
        __wbg_new_a70fbab9066b301f: function() {
            const ret = new Array();
            return ret;
        },
        __wbg_new_ab79df5bd7c26067: function() {
            const ret = new Object();
            return ret;
        },
        __wbg_next_11b99ee6237339e3: function() { return handleError(function (arg0) {
            const ret = arg0.next();
            return ret;
        }, arguments); },
        __wbg_next_e01a967809d1aa68: function(arg0) {
            const ret = arg0.next;
            return ret;
        },
        __wbg_prototypesetcall_d62e5099504357e6: function(arg0, arg1, arg2) {
            Uint8Array.prototype.set.call(getArrayU8FromWasm0(arg0, arg1), arg2);
        },
        __wbg_set_282384002438957f: function(arg0, arg1, arg2) {
            arg0[arg1 >>> 0] = arg2;
        },
        __wbg_set_6be42768c690e380: function(arg0, arg1, arg2) {
            arg0[arg1] = arg2;
        },
        __wbg_set_bf7251625df30a02: function(arg0, arg1, arg2) {
            const ret = arg0.set(arg1, arg2);
            return ret;
        },
        __wbg_stack_3b0d974bbf31e44f: function(arg0, arg1) {
            const ret = arg1.stack;
            const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
            getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
        },
        __wbg_value_21fc78aab0322612: function(arg0) {
            const ret = arg0.value;
            return ret;
        },
        __wbindgen_cast_0000000000000001: function(arg0) {
            // Cast intrinsic for `F64 -> Externref`.
            const ret = arg0;
            return ret;
        },
        __wbindgen_cast_0000000000000002: function(arg0) {
            // Cast intrinsic for `I64 -> Externref`.
            const ret = arg0;
            return ret;
        },
        __wbindgen_cast_0000000000000003: function(arg0, arg1) {
            // Cast intrinsic for `Ref(String) -> Externref`.
            const ret = getStringFromWasm0(arg0, arg1);
            return ret;
        },
        __wbindgen_cast_0000000000000004: function(arg0) {
            // Cast intrinsic for `U64 -> Externref`.
            const ret = BigInt.asUintN(64, arg0);
            return ret;
        },
        __wbindgen_init_externref_table: function() {
            const table = wasm.__wbindgen_externrefs;
            const offset = table.grow(4);
            table.set(0, undefined);
            table.set(offset + 0, undefined);
            table.set(offset + 1, null);
            table.set(offset + 2, true);
            table.set(offset + 3, false);
        },
    };
    return {
        __proto__: null,
        "./oxidocs_wasm_bg.js": import0,
    };
}

const DocxDocumentFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_docxdocument_free(ptr >>> 0, 1));

function addToExternrefTable0(obj) {
    const idx = wasm.__externref_table_alloc();
    wasm.__wbindgen_externrefs.set(idx, obj);
    return idx;
}

function debugString(val) {
    // primitive types
    const type = typeof val;
    if (type == 'number' || type == 'boolean' || val == null) {
        return  `${val}`;
    }
    if (type == 'string') {
        return `"${val}"`;
    }
    if (type == 'symbol') {
        const description = val.description;
        if (description == null) {
            return 'Symbol';
        } else {
            return `Symbol(${description})`;
        }
    }
    if (type == 'function') {
        const name = val.name;
        if (typeof name == 'string' && name.length > 0) {
            return `Function(${name})`;
        } else {
            return 'Function';
        }
    }
    // objects
    if (Array.isArray(val)) {
        const length = val.length;
        let debug = '[';
        if (length > 0) {
            debug += debugString(val[0]);
        }
        for(let i = 1; i < length; i++) {
            debug += ', ' + debugString(val[i]);
        }
        debug += ']';
        return debug;
    }
    // Test for built-in
    const builtInMatches = /\[object ([^\]]+)\]/.exec(toString.call(val));
    let className;
    if (builtInMatches && builtInMatches.length > 1) {
        className = builtInMatches[1];
    } else {
        // Failed to match the standard '[object ClassName]'
        return toString.call(val);
    }
    if (className == 'Object') {
        // we're a user defined class or Object
        // JSON.stringify avoids problems with cycles, and is generally much
        // easier than looping through ownProperties of `val`.
        try {
            return 'Object(' + JSON.stringify(val) + ')';
        } catch (_) {
            return 'Object';
        }
    }
    // errors
    if (val instanceof Error) {
        return `${val.name}: ${val.message}\n${val.stack}`;
    }
    // TODO we could test for more things here, like `Set`s and `Map`s.
    return className;
}

function getArrayU8FromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    return getUint8ArrayMemory0().subarray(ptr / 1, ptr / 1 + len);
}

let cachedDataViewMemory0 = null;
function getDataViewMemory0() {
    if (cachedDataViewMemory0 === null || cachedDataViewMemory0.buffer.detached === true || (cachedDataViewMemory0.buffer.detached === undefined && cachedDataViewMemory0.buffer !== wasm.memory.buffer)) {
        cachedDataViewMemory0 = new DataView(wasm.memory.buffer);
    }
    return cachedDataViewMemory0;
}

let cachedFloat32ArrayMemory0 = null;
function getFloat32ArrayMemory0() {
    if (cachedFloat32ArrayMemory0 === null || cachedFloat32ArrayMemory0.byteLength === 0) {
        cachedFloat32ArrayMemory0 = new Float32Array(wasm.memory.buffer);
    }
    return cachedFloat32ArrayMemory0;
}

function getStringFromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    return decodeText(ptr, len);
}

let cachedUint8ArrayMemory0 = null;
function getUint8ArrayMemory0() {
    if (cachedUint8ArrayMemory0 === null || cachedUint8ArrayMemory0.byteLength === 0) {
        cachedUint8ArrayMemory0 = new Uint8Array(wasm.memory.buffer);
    }
    return cachedUint8ArrayMemory0;
}

function handleError(f, args) {
    try {
        return f.apply(this, args);
    } catch (e) {
        const idx = addToExternrefTable0(e);
        wasm.__wbindgen_exn_store(idx);
    }
}

function isLikeNone(x) {
    return x === undefined || x === null;
}

function passArray8ToWasm0(arg, malloc) {
    const ptr = malloc(arg.length * 1, 1) >>> 0;
    getUint8ArrayMemory0().set(arg, ptr / 1);
    WASM_VECTOR_LEN = arg.length;
    return ptr;
}

function passArrayF32ToWasm0(arg, malloc) {
    const ptr = malloc(arg.length * 4, 4) >>> 0;
    getFloat32ArrayMemory0().set(arg, ptr / 4);
    WASM_VECTOR_LEN = arg.length;
    return ptr;
}

function passStringToWasm0(arg, malloc, realloc) {
    if (realloc === undefined) {
        const buf = cachedTextEncoder.encode(arg);
        const ptr = malloc(buf.length, 1) >>> 0;
        getUint8ArrayMemory0().subarray(ptr, ptr + buf.length).set(buf);
        WASM_VECTOR_LEN = buf.length;
        return ptr;
    }

    let len = arg.length;
    let ptr = malloc(len, 1) >>> 0;

    const mem = getUint8ArrayMemory0();

    let offset = 0;

    for (; offset < len; offset++) {
        const code = arg.charCodeAt(offset);
        if (code > 0x7F) break;
        mem[ptr + offset] = code;
    }
    if (offset !== len) {
        if (offset !== 0) {
            arg = arg.slice(offset);
        }
        ptr = realloc(ptr, len, len = offset + arg.length * 3, 1) >>> 0;
        const view = getUint8ArrayMemory0().subarray(ptr + offset, ptr + len);
        const ret = cachedTextEncoder.encodeInto(arg, view);

        offset += ret.written;
        ptr = realloc(ptr, len, offset, 1) >>> 0;
    }

    WASM_VECTOR_LEN = offset;
    return ptr;
}

function takeFromExternrefTable0(idx) {
    const value = wasm.__wbindgen_externrefs.get(idx);
    wasm.__externref_table_dealloc(idx);
    return value;
}

let cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
cachedTextDecoder.decode();
const MAX_SAFARI_DECODE_BYTES = 2146435072;
let numBytesDecoded = 0;
function decodeText(ptr, len) {
    numBytesDecoded += len;
    if (numBytesDecoded >= MAX_SAFARI_DECODE_BYTES) {
        cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
        cachedTextDecoder.decode();
        numBytesDecoded = len;
    }
    return cachedTextDecoder.decode(getUint8ArrayMemory0().subarray(ptr, ptr + len));
}

const cachedTextEncoder = new TextEncoder();

if (!('encodeInto' in cachedTextEncoder)) {
    cachedTextEncoder.encodeInto = function (arg, view) {
        const buf = cachedTextEncoder.encode(arg);
        view.set(buf);
        return {
            read: arg.length,
            written: buf.length
        };
    };
}

let WASM_VECTOR_LEN = 0;

let wasmModule, wasm;
function __wbg_finalize_init(instance, module) {
    wasm = instance.exports;
    wasmModule = module;
    cachedDataViewMemory0 = null;
    cachedFloat32ArrayMemory0 = null;
    cachedUint8ArrayMemory0 = null;
    wasm.__wbindgen_start();
    return wasm;
}

async function __wbg_load(module, imports) {
    if (typeof Response === 'function' && module instanceof Response) {
        if (typeof WebAssembly.instantiateStreaming === 'function') {
            try {
                return await WebAssembly.instantiateStreaming(module, imports);
            } catch (e) {
                const validResponse = module.ok && expectedResponseType(module.type);

                if (validResponse && module.headers.get('Content-Type') !== 'application/wasm') {
                    console.warn("`WebAssembly.instantiateStreaming` failed because your server does not serve Wasm with `application/wasm` MIME type. Falling back to `WebAssembly.instantiate` which is slower. Original error:\n", e);

                } else { throw e; }
            }
        }

        const bytes = await module.arrayBuffer();
        return await WebAssembly.instantiate(bytes, imports);
    } else {
        const instance = await WebAssembly.instantiate(module, imports);

        if (instance instanceof WebAssembly.Instance) {
            return { instance, module };
        } else {
            return instance;
        }
    }

    function expectedResponseType(type) {
        switch (type) {
            case 'basic': case 'cors': case 'default': return true;
        }
        return false;
    }
}

function initSync(module) {
    if (wasm !== undefined) return wasm;


    if (module !== undefined) {
        if (Object.getPrototypeOf(module) === Object.prototype) {
            ({module} = module)
        } else {
            console.warn('using deprecated parameters for `initSync()`; pass a single object instead')
        }
    }

    const imports = __wbg_get_imports();
    if (!(module instanceof WebAssembly.Module)) {
        module = new WebAssembly.Module(module);
    }
    const instance = new WebAssembly.Instance(module, imports);
    return __wbg_finalize_init(instance, module);
}

async function __wbg_init(module_or_path) {
    if (wasm !== undefined) return wasm;


    if (module_or_path !== undefined) {
        if (Object.getPrototypeOf(module_or_path) === Object.prototype) {
            ({module_or_path} = module_or_path)
        } else {
            console.warn('using deprecated parameters for the initialization function; pass a single object instead')
        }
    }

    if (module_or_path === undefined) {
        module_or_path = new URL('oxidocs_wasm_bg.wasm', import.meta.url);
    }
    const imports = __wbg_get_imports();

    if (typeof module_or_path === 'string' || (typeof Request === 'function' && module_or_path instanceof Request) || (typeof URL === 'function' && module_or_path instanceof URL)) {
        module_or_path = fetch(module_or_path);
    }

    const { instance, module } = await __wbg_load(await module_or_path, imports);

    return __wbg_finalize_init(instance, module);
}

export { initSync, __wbg_init as default };

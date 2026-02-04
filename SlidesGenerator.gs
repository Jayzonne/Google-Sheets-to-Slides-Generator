/**
 * @file SlidesGenerator.gs
 * Generates a new Google Slides file from a template:
 * - 1 slide per checked database row
 * - Replaces {{header}} placeholders
 * - Supports images using placeholders {{FIELD}} in the template slide
 *   and IMAGE_N_* config blocks (FIELD, SOURCE, FIT)
 * - Keeps the placeholder shape (style, border, shadows, etc.)
 * - Inserts the image at the EXACT same position/size/rotation as the placeholder
 * - Places the image at the SAME "layer zone" by positioning it just behind the
 *   placeholder's top-level anchor (shape or group)
 */

 /**
  * @typedef {Object} GenerationResult
  * @property {string} fileId
  * @property {string} fileName
  * @property {number} slidesGenerated
  */

class SlidesGenerator {
  /**
   * @param {AppConfig} config
   */
  constructor(config) {
    this.config = config;
  }

  /**
   * @param {{headers: string[], rows: Array<Array<*>>}} dataset
   * @return {GenerationResult}
   */
  generate(dataset) {
    const templateFile = DriveApp.getFileById(this.config.templateSlidesId);
    const outputFolder = DriveApp.getFolderById(this.config.outputFolderId);

    const outputName =
      Utils.applyFileNameTokens_(this.config.outputFileNamePattern) || 'Generated Slides';

    const newFile = templateFile.makeCopy(outputName, outputFolder);
    const presentation = SlidesApp.openById(newFile.getId());

    const slides = presentation.getSlides();
    if (slides.length < this.config.templateSlideIndex) {
      throw new Error(
        `Template slide index (${this.config.templateSlideIndex}) is out of range. Template has ${slides.length} slides.`
      );
    }

    const templateSlide = slides[this.config.templateSlideIndex - 1];

    // Clean-deck strategy: keep the template slide only, remove others.
    for (let i = slides.length - 1; i >= 0; i--) {
      if (i !== this.config.templateSlideIndex - 1) slides[i].remove();
    }

    const imagesCfg = Array.isArray(this.config.images) ? this.config.images : [];
    const imageFields = new Set(imagesCfg.map((img) => String(img.field || '').trim()).filter(Boolean));

    let generatedCount = 0;

    for (const row of dataset.rows) {
      if (Utils.isEmptyRow_(row)) continue;

      const rowData = this._buildRowData_(dataset.headers, row);

      // appendSlide preserves sheet row order
      const newSlide = presentation.appendSlide(templateSlide);

      // 1) Replace text placeholders (excluding image fields so the token stays in the slide)
      const replacements = this._buildReplacementsExcluding_(rowData, imageFields);
      this._replaceInSlide_(newSlide, replacements);

      // 2) Apply images
      if (imagesCfg.length) {
        this._applyImages_(newSlide, rowData, imagesCfg);
      }

      generatedCount++;
    }

    // Remove the original template slide (used as blueprint)
    templateSlide.remove();

    presentation.saveAndClose();

    return {
      fileId: newFile.getId(),
      fileName: newFile.getName(),
      slidesGenerated: generatedCount,
    };
  }

  /**
   * Builds a header->string map from headers + a row.
   * @param {string[]} headers
   * @param {Array<*>} row
   * @return {Object<string,string>}
   * @private
   */
  _buildRowData_(headers, row) {
    const map = {};
    headers.forEach((h, i) => {
      const key = String(h).trim();
      const value = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
      map[key] = value;
    });
    return map;
  }

  /**
   * Builds a token->value map ({{header}} -> value), excluding some fields.
   * @param {Object<string,string>} rowData
   * @param {Set<string>} excludeFields
   * @return {Object<string,string>}
   * @private
   */
  _buildReplacementsExcluding_(rowData, excludeFields) {
    const map = {};
    Object.keys(rowData).forEach((field) => {
      if (excludeFields && excludeFields.has(field)) return;
      map[`{{${field}}}`] = rowData[field];
    });
    return map;
  }

  /**
   * Replaces tokens in a slide.
   * @param {GoogleAppsScript.Slides.Slide} slide
   * @param {Object<string,string>} replacements
   * @private
   */
  _replaceInSlide_(slide, replacements) {
    Object.keys(replacements).forEach((token) => {
      slide.replaceAllText(token, replacements[token]);
    });
  }

  /**
   * Apply all configured images on a slide.
   *
   * Rules:
   * - Placeholder token is always {{FIELD}}.
   * - We KEEP the placeholder shape (so its formatting/border stays).
   * - We remove only the token text inside it (preserving text formatting).
   * - The inserted image copies position/size/rotation from the placeholder shape.
   * - Z-order: we place the image just behind the placeholder's TOP-LEVEL anchor.
   *
   * Notes:
   * - If your placeholder is inside a GROUP, we can only position relative to the group
   *   (because the image is inserted as a top-level element). For best results, avoid
   *   grouping the placeholder if you need perfect layering.
   *
   * @param {GoogleAppsScript.Slides.Slide} slide
   * @param {Object<string,string>} rowData
   * @param {Array<{field:string, source:string, fit:string}>} imagesCfg
   * @private
   */
  _applyImages_(slide, rowData, imagesCfg) {
    for (const cfg of imagesCfg) {
      const field = String(cfg.field || '').trim();
      if (!field) continue;

      const token = `{{${field}}}`;
      const raw = String(rowData[field] || '').trim();

      // Find placeholder shapes that contain token; includes grouped shapes.
      // Returns items: { shape, anchorId }
      const targets = this._findPlaceholderShapes_(slide, token);

      if (!targets.length) continue;

      // If no image value: just remove token text (keep formatting/shape)
      if (!raw) {
        for (const t of targets) {
          this._removeTokenKeepFormatting_(t.shape, token);
        }
        continue;
      }

      const blob = this._resolveImageBlob_(cfg, raw);

      for (const t of targets) {
        const shape = t.shape;

        const left = shape.getLeft();
        const top = shape.getTop();
        const width = shape.getWidth();
        const height = shape.getHeight();
        const rotation = shape.getRotation();

        // Insert image
        const img = slide.insertImage(blob);

        // Apply rotation + geometry
        img.setLeft(left).setTop(top);

        const fit = String(cfg.fit || 'CONTAIN').trim().toUpperCase();

        if (fit === 'STRETCH') {
          img.setWidth(width).setHeight(height);
        } else if (fit === 'COVER') {
          // "Cover" effect: fill the box; best-effort crop by using replace(blob, true)
          // after setting the box size.
          img.setWidth(width).setHeight(height);
          try {
            // cropToFit = true (Apps Script Slides Image.replace supports this on most accounts)
            img.replace(blob, true /** cropToFit */ , true);
          } catch (e) {
            // If not supported in this runtime, fallback to stretch (still same size)
          }
        } else {
          // CONTAIN (default): keep ratio inside box, center
          const iw = img.getWidth();
          const ih = img.getHeight();
          if (iw && ih) {
            const scale = Math.min(width / iw, height / ih);
            const newW = iw * scale;
            const newH = ih * scale;
            img.setWidth(newW).setHeight(newH);
            img.setLeft(left + (width - newW) / 2);
            img.setTop(top + (height - newH) / 2);
          } else {
            img.setWidth(width).setHeight(height);
          }
        }

        // Rotation last (so it matches placeholder exactly)
        img.setRotation(rotation);

        // Put the image in the correct "layer zone"
        // => just behind the placeholder's top-level anchor
        this._placeImageJustBehindAnchor_(slide, img, t.anchorId);

        // Remove token text from placeholder but keep its style
        this._removeTokenKeepFormatting_(shape, token);
      }
    }
  }

  /**
   * Find shapes that contain the token (even inside groups).
   * Returns {shape, anchorId} where anchorId is the top-level element id
   * (group if grouped, else the shape itself).
   *
   * Safe against shapes that throw on getText().
   *
   * @param {GoogleAppsScript.Slides.Slide} slide
   * @param {string} token
   * @return {Array<{shape: GoogleAppsScript.Slides.Shape, anchorId: string}>}
   * @private
   */
  _findPlaceholderShapes_(slide, token) {
    /** @type {Array<{shape: GoogleAppsScript.Slides.Shape, anchorId: string}>} */
    const out = [];

    const walk = (elements) => {
      for (const el of elements) {
        const type = el.getPageElementType();

        if (type === SlidesApp.PageElementType.GROUP) {
          walk(el.asGroup().getChildren());
          continue;
        }

        if (type !== SlidesApp.PageElementType.SHAPE) continue;

        const shape = el.asShape();

        let txt = '';
        try {
          txt = shape.getText().asString();
        } catch (e) {
          continue; // shape not text-capable
        }

        if (txt && txt.indexOf(token) !== -1) {
          const anchor = this._getTopLevelAnchor_(shape);
          out.push({ shape, anchorId: anchor.getObjectId() });
        }
      }
    };

    walk(slide.getPageElements());
    return out;
  }

  /**
   * Returns the top-level anchor for a page element:
   * - if element is in a group, returns the highest parent group
   * - otherwise returns the element itself
   *
   * @param {GoogleAppsScript.Slides.PageElement} el
   * @return {GoogleAppsScript.Slides.PageElement}
   * @private
   */
  _getTopLevelAnchor_(el) {
    let current = el;
    try {
      while (current.getParentGroup && current.getParentGroup()) {
        current = current.getParentGroup();
      }
    } catch (e) {
      // ignore
    }
    return current;
  }

  /**
   * Removes token from shape text while keeping formatting.
   * If the text becomes empty, we append a zero-width char to avoid any autofit weirdness,
   * without changing visible formatting.
   *
   * @param {GoogleAppsScript.Slides.Shape} shape
   * @param {string} token
   * @private
   */
  _removeTokenKeepFormatting_(shape, token) {
    try {
      const tr = shape.getText();
      tr.replaceAllText(token, '', false);

      // If text becomes empty/whitespace, keep a zero-width marker to preserve box behavior
      const after = String(tr.asString() || '');
      const cleaned = after.replace(/[\s\u200B]+/g, '');
      if (!cleaned) {
        try {
          // Only append if truly empty (avoid accumulating chars)
          tr.appendText('\u200B');
        } catch (e2) {
          // ignore
        }
      }
    } catch (e) {
      // ignore
    }
  }

  /**
   * Places the image just behind the anchor (same "frontness" zone as anchor).
   * We do it reliably by:
   * - inferring the order direction of slide.getPageElements()
   * - bringing image to front
   * - sending backward until it is exactly one step behind anchor
   *
   * @param {GoogleAppsScript.Slides.Slide} slide
   * @param {GoogleAppsScript.Slides.Image} img
   * @param {string} anchorId
   * @private
   */
  _placeImageJustBehindAnchor_(slide, img, anchorId) {
    try {
      // Infer ordering (back-to-front vs front-to-back)
      const direction = this._inferZOrderDirection_(slide, img);

      // Start from front so we can walk backwards deterministically
      img.bringToFront();

      const maxSteps = Math.max(10, slide.getPageElements().length + 10);

      for (let i = 0; i < maxSteps; i++) {
        const elements = slide.getPageElements();
        const idxAnchor = elements.findIndex((e) => e.getObjectId() === anchorId);
        const idxImg = elements.findIndex((e) => e.getObjectId() === img.getObjectId());

        if (idxAnchor === -1 || idxImg === -1) break;

        const desired =
          direction === 'BACK_TO_FRONT'
            ? Math.max(0, idxAnchor - 1)                 // one behind (smaller index)
            : Math.min(elements.length - 1, idxAnchor + 1); // one behind (larger index)

        if (idxImg === desired) break;

        img.sendBackward();
      }
    } catch (e) {
      // If anything fails (groups / API limitation), we keep default insertion order.
    }
  }

  /**
   * Infers whether slide.getPageElements() is ordered back->front or front->back
   * by moving the image to front and to back and comparing indices.
   *
   * Returns:
   * - 'BACK_TO_FRONT' if "front" index > "back" index
   * - 'FRONT_TO_BACK' otherwise
   *
   * @param {GoogleAppsScript.Slides.Slide} slide
   * @param {GoogleAppsScript.Slides.PageElement} el
   * @return {'BACK_TO_FRONT'|'FRONT_TO_BACK'}
   * @private
   */
  _inferZOrderDirection_(slide, el) {
    try {
      el.bringToFront();
      let idxFront = slide.getPageElements().findIndex((e) => e.getObjectId() === el.getObjectId());

      el.sendToBack();
      let idxBack = slide.getPageElements().findIndex((e) => e.getObjectId() === el.getObjectId());

      // Put it back to front so our placement routine starts from a known state
      el.bringToFront();

      return idxFront > idxBack ? 'BACK_TO_FRONT' : 'FRONT_TO_BACK';
    } catch (e) {
      // Default assumption (most accounts behave like this)
      return 'BACK_TO_FRONT';
    }
  }

  /**
   * Resolve blob for image.
   * - DRIVE_ID: value is drive file ID OR a drive/docs URL containing an ID
   * - URL: value is a direct image URL
   *
   * @param {{field:string, source:string}} cfg
   * @param {string} rawValue
   * @return {GoogleAppsScript.Base.Blob}
   * @private
   */
  _resolveImageBlob_(cfg, rawValue) {
    const source = String(cfg.source || 'DRIVE_ID').trim().toUpperCase();

    if (source === 'URL') {
      const res = UrlFetchApp.fetch(rawValue, { muteHttpExceptions: true, followRedirects: true });
      const code = res.getResponseCode();
      if (code < 200 || code >= 300) {
        throw new Error(`Failed to fetch image URL (HTTP ${code}) for field "${cfg.field}".`);
      }
      return res.getBlob();
    }

    // Default DRIVE_ID
    const id = this._extractDriveId_(rawValue) || rawValue;
    return DriveApp.getFileById(id).getBlob();
  }

  /**
   * Extract Drive file ID from URL or string if possible.
   * @param {string} s
   * @return {string|null}
   * @private
   */
  _extractDriveId_(s) {
    const str = String(s || '').trim();
    if (!str) return null;

    // - https://drive.google.com/file/d/<ID>/view
    // - https://drive.google.com/open?id=<ID>
    // - https://docs.google.com/.../d/<ID>/edit
    const m1 = /\/d\/([a-zA-Z0-9_-]{20,})/.exec(str);
    if (m1) return m1[1];

    const m2 = /[?&]id=([a-zA-Z0-9_-]{20,})/.exec(str);
    if (m2) return m2[1];

    // Fallback: if the string itself looks like an ID
    const m3 = /^([a-zA-Z0-9_-]{20,})$/.exec(str);
    if (m3) return m3[1];

    return null;
  }
}

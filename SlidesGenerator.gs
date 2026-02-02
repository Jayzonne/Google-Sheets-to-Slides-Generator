/**
 * @file SlidesGenerator.gs
 * Generates a new Google Slides file from a template:
 * - 1 slide per database row
 * - Replaces {{header}} placeholders
 * - Preserves row order using appendSlide()
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

    let generatedCount = 0;

    for (const row of dataset.rows) {
      if (Utils.isEmptyRow_(row)) continue;

      const replacements = this._buildReplacements_(dataset.headers, row);

      // ✅ appendSlide preserves sheet row order
      const newSlide = presentation.appendSlide(templateSlide);
      this._replaceInSlide_(newSlide, replacements);

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

  // "Private" by convention
  _buildReplacements_(headers, row) {
    const map = {};
    headers.forEach((h, i) => {
      const key = String(h).trim();
      const value = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
      map[`{{${key}}}`] = value;
    });
    return map;
  }

  // "Private" by convention
  _replaceInSlide_(slide, replacements) {
    Object.keys(replacements).forEach((token) => {
      slide.replaceAllText(token, replacements[token]);
    });
  }
}

const CODEBOOK_HEADER_CODE = 'Code';
const CODEBOOK_HEADER_TYPE = 'Type';
const CODEBOOK_TYPE_CODE = 'code';
const CODEBOOK_TYPE_FLAG = 'flag';
const CODEBOOK_PATTERN = /(\w+)_codebook/;
const CODING_PATTERN = /(\w+)_codes(_\w+)?/;
const FINAL_CODES_PATTERN = /(\w+)_codes_final/;

const HEADER_ROW = 1;
const FIRST_ROW = 2; // Assuming a header, row 2 is always the first row.

const CODES_SEPARATOR = ',';

const CODEBOOK_HEADER_FINAL = 'Code - final';
const CODEBOOK_SHEET_NAME = (questionId: string) => questionId + '_codebook';

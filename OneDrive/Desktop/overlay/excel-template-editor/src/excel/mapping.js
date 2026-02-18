export const matterDetailsFields = [
  { key: 'ebayDocketNo', label: 'eBay Docket No.', a1: 'B5', type: 'text', required: true },
  { key: 'ocDocketNo', label: 'OC Docket No.', a1: 'D5', type: 'text' },
  { key: 'ebayAttorney', label: 'eBay Attorney', a1: 'B6', type: 'text' },
  { key: 'ocAttorney', label: 'OC Attorney', a1: 'D6', type: 'text' },
  { key: 'applicationNo', label: 'Application No.', a1: 'B7', type: 'text' },
  { key: 'filingDate', label: 'Filing Date', a1: 'D7', type: 'date' },
  { key: 'applicationType', label: 'Application Type', a1: 'B8', type: 'text' },
  { key: 'priorityDate', label: 'Priority Date', a1: 'D8', type: 'date' },
  { key: 'portfolioCategory', label: 'Portfolio Category', a1: 'B9', type: 'text' },
  { key: 'foreignAssociate', label: 'Foreign Associate', a1: 'D9', type: 'text' },
  { key: 'disclosureNo', label: 'Disclosure No.', a1: 'B10', type: 'text' },
  { key: 'estYearsToExpiration', label: 'Est. Years to Expiration', a1: 'D10', type: 'number' },
]

export const inventorFields = [
  { key: 'fullLegalName', label: 'Full Legal Name', offset: 1 },
  { key: 'email', label: 'Email', offset: 3 },
  { key: 'residence', label: 'Residence', offset: 4 },
  { key: 'mailing', label: 'Mailing', offset: 7 },
  { key: 'citizenship', label: 'Citizenship', offset: 10 },
  { key: 'conception', label: 'Conception', offset: 12 },
  { key: 'absences', label: 'Absences', offset: 15 },
]

export function inventorStartRow(indexZeroBased) {
  return 3 + indexZeroBased * 17
}

export const actionStages = [
  { key: 'ENGAGEMENT', label: 'Engagement', column: 'B' },
  { key: 'DISCLOSURE_PRE', label: 'Disclosure: Pre-Search', column: 'C' },
  { key: 'DISCLOSURE_POST', label: 'Disclosure: Post Search', column: 'D' },
  { key: 'RESTRICTION', label: 'Restriction', column: 'E' },
  { key: 'OFFICE_ACTION', label: 'Office Action / Appeal Decision', column: 'F' },
  { key: 'ADVISORY_AFCP', label: "Advisory/AFCP/Pre-Appeal/Examiner's Answer", column: 'G' },
  { key: 'NOA', label: 'Notice of Allowance', column: 'H' },
]

export const actionsSections = [
  { key: 'metadata', label: 'Event Metadata', startRow: 2, endRow: 4 },
  { key: 'strategy', label: 'Main Strategy', startRow: 6, endRow: 17 },
  { key: 'continuation', label: 'Continuation Proposal', startRow: 18, endRow: 28 },
]

export const triadRows = {
  main: [9, 10, 11],
  continuation: [25, 26, 27],
}


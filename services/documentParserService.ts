import mammoth from 'mammoth';
import * as pdfjs from 'pdfjs-dist';

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.min.mjs',
  import.meta.url
).toString();

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

type DocumentKind = 'pdf' | 'docx' | 'text' | 'unsupported';

const detectDocumentKind = (file: File): DocumentKind => {
  const name = file.name.toLowerCase();

  if (file.type === 'application/pdf' || name.endsWith('.pdf')) return 'pdf';
  if (file.type === DOCX_MIME || name.endsWith('.docx')) return 'docx';
  if (
    file.type.startsWith('text/') ||
    name.endsWith('.txt') ||
    name.endsWith('.md')
  ) {
    return 'text';
  }

  return 'unsupported';
};

const extractPdfText = async (file: File): Promise<string> => {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: buffer }).promise;
  const pages: string[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    pages.push(
      content.items
        .map((item) => ('str' in item ? item.str : ''))
        .join(' ')
        .trim()
    );
  }

  return pages.filter(Boolean).join('\n\n');
};

const extractDocxText = async (file: File): Promise<string> => {
  const buffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value.trim();
};

export const extractDocumentText = async (file: File): Promise<string> => {
  const kind = detectDocumentKind(file);

  switch (kind) {
    case 'text':
      return file.text();
    case 'docx':
      return extractDocxText(file);
    case 'pdf': {
      const text = await extractPdfText(file);
      if (!text.trim()) {
        throw new Error('No extractable text found (scanned/image PDF?)');
      }
      return text;
    }
    default:
      throw new Error(`Unsupported file type: ${file.name}`);
  }
};

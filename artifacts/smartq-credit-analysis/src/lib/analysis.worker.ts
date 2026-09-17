import { analyzeWorkbook, type AnalysisResult } from './analysis';

type WorkerRequest = {
  input: ArrayBuffer | string;
  format: 'excel' | 'csv';
  sourceName: string;
};

type WorkerResponse =
  | { type: 'success'; result: AnalysisResult }
  | { type: 'error'; message: string; missingColumns: string[] };

self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  try {
    const result = await analyzeWorkbook(event.data.input, event.data.format, event.data.sourceName);
    self.postMessage({ type: 'success', result } satisfies WorkerResponse);
  } catch (error) {
    const analysisError = error as Error & { missingColumns?: string[] };
    self.postMessage({
      type: 'error',
      message: analysisError.message || 'The workbook could not be read. It may be damaged or protected.',
      missingColumns: analysisError.missingColumns ?? [],
    } satisfies WorkerResponse);
  }
};

export {};

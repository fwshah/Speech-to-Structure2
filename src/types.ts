export interface TranscriptionSegment {
  speaker: string;
  text: string;
  startTime: number; // in seconds
  endTime: number; // in seconds
}

export interface TranscriptionReport {
  title: string;
  summary: string;
  segments: TranscriptionSegment[];
  speakers: string[];
}

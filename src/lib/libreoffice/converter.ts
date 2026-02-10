/**
 * LibreOffice WASM Converter
 * 
 * Uses @matbee/libreoffice-converter for document conversion.
 */

import { WorkerBrowserConverter } from '@matbee/libreoffice-converter/browser';

const LIBREOFFICE_PATH = '/libreoffice-wasm/';

export interface LoadProgress {
    phase: 'loading' | 'initializing' | 'converting' | 'complete' | 'ready';
    percent: number;
    message: string;
}

export type ProgressCallback = (progress: LoadProgress) => void;

let converterInstance: LibreOfficeConverter | null = null;

export class LibreOfficeConverter {
    private converter: WorkerBrowserConverter | null = null;
    private initialized = false;
    private initializing = false;
    private basePath: string;

    constructor(basePath?: string) {
        this.basePath = basePath || LIBREOFFICE_PATH;
    }

    async initialize(onProgress?: ProgressCallback): Promise<void> {
        if (this.initialized) return;

        if (this.initializing) {
            while (this.initializing) {
                await new Promise(r => setTimeout(r, 100));
            }
            return;
        }

        this.initializing = true;
        let progressCallback = onProgress;

        try {
            progressCallback?.({ phase: 'loading', percent: 0, message: 'Loading conversion engine...' });

            this.converter = new WorkerBrowserConverter({
                sofficeJs: `${this.basePath}soffice.js`,
                sofficeWasm: `${this.basePath}soffice.wasm`,
                sofficeData: `${this.basePath}soffice.data`,
                sofficeWorkerJs: `${this.basePath}soffice.worker.js`,
                browserWorkerJs: `${this.basePath}browser.worker.global.js`,
                verbose: false,
                onProgress: (info: { phase: string; percent: number; message: string }) => {
                    if (progressCallback && !this.initialized) {
                        progressCallback({
                            phase: info.phase as LoadProgress['phase'],
                            percent: info.percent,
                            message: `Loading conversion engine (${Math.round(info.percent)}%)...`
                        });
                    }
                },
                onReady: () => {
                    console.log('[LibreOffice] Ready!');
                },
                onError: (error: Error) => {
                    console.error('[LibreOffice] Error:', error);
                },
            });

            await this.converter.initialize();
            this.initialized = true;
            progressCallback?.({ phase: 'ready', percent: 100, message: 'Conversion engine ready!' });
            progressCallback = undefined;
        } finally {
            this.initializing = false;
        }
    }

    isReady(): boolean {
        return this.initialized && this.converter !== null;
    }

    async convert(file: File, outputFormat: string): Promise<Blob> {
        if (!this.converter) {
            throw new Error('Converter not initialized');
        }

        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const ext = file.name.split('.').pop()?.toLowerCase() || '';

        const result = await this.converter.convert(uint8Array, {
            outputFormat: outputFormat as any,
            inputFormat: ext as any,
        }, file.name);

        const data = new Uint8Array(result.data);
        return new Blob([data], { type: result.mimeType });
    }

    async convertToPdf(file: File): Promise<Blob> {
        return this.convert(file, 'pdf');
    }

    async destroy(): Promise<void> {
        if (this.converter) {
            await this.converter.destroy();
        }
        this.converter = null;
        this.initialized = false;
    }
}

export function getLibreOfficeConverter(basePath?: string): LibreOfficeConverter {
    if (!converterInstance) {
        converterInstance = new LibreOfficeConverter(basePath);
    }
    return converterInstance;
}

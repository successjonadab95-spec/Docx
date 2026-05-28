/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState } from 'react';
import * as docx from 'docx';
import { saveAs } from 'file-saver';
import { Buffer } from 'buffer';
import { 
  Play, Code2, Terminal, AlertCircle, FileText, 
  Trash2, ClipboardPaste, CheckCircle2, Loader2 
} from 'lucide-react';
import Editor from 'react-simple-code-editor';
import Prism from 'prismjs';
import 'prismjs/components/prism-javascript';
import 'prismjs/themes/prism-tomorrow.css';

// Polyfills required for docx Packer.toBuffer() in browser
(window as any).Buffer = Buffer;
(window as any).global = window;

const DEFAULT_SCRIPT = `// Paste your script here...
`;

export default function App() {
  const [code, setCode] = useState(DEFAULT_SCRIPT);
  const [isRunning, setIsRunning] = useState(false);
  const [logs, setLogs] = useState<{ type: 'info' | 'error' | 'success', text: string }[]>([
    { type: 'info', text: 'System ready. Waiting for script execution...' }
  ]);

  const addLog = (type: 'info' | 'error' | 'success', text: string) => {
    setLogs(prev => [...prev, { type, text }]);
  };

  const handleClear = () => setCode('');
  
  const handlePaste = async () => {
    try {
      const text = await navigator.clipboard.readText();
      setCode(text);
    } catch (err) {
      addLog('error', 'Failed to read from clipboard. Please paste manually.');
    }
  };

  const runScript = async () => {
    setIsRunning(true);
    setLogs([]);

    try {
      // Mock require function to hijack 'docx' and 'fs'
      const customRequire = (moduleName: string) => {
        if (moduleName.includes('docx')) {
          return {
            ...docx,
            Packer: {
              ...docx.Packer,
              toBuffer: async (doc: any) => {
                // Intercept toBuffer and use toBlob instead because jszip
                // in the browser bundle drops nodebuffer support.
                const blob = await docx.Packer.toBlob(doc);
                const arrayBuffer = await blob.arrayBuffer();
                return Buffer.from(arrayBuffer);
              }
            }
          };
        }
        if (moduleName === 'fs') {
          return {
            writeFileSync: (filePath: string, data: any) => {
              const filename = filePath.split('/').pop() || 'document.docx';
              try {
                // data is ideally a Buffer from Packer.toBuffer
                const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                saveAs(blob, filename);
                addLog('success', `Download triggered successfully: ${filename}`);
              } catch (e: any) {
                addLog('error', `Failed to save file: ${e.message}`);
              } finally {
                setIsRunning(false);
              }
            }
          };
        }
        return {}; // Return empty object for unsupported requires
      };

      // Mock console to capture user script logs
      const customConsole = {
        log: (...args: any[]) => {
          console.log(...args);
          addLog('info', args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '));
        },
        error: (...args: any[]) => {
          console.error(...args);
          addLog('error', args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '));
        }
      };

      addLog('info', 'Compiling and executing script...');

      // Create a function from the user's code
      const fn = new Function('require', 'console', 'Buffer', code);
      
      // Execute the user's code, passing in our hijacked dependencies
      fn(customRequire, customConsole, Buffer);
      
      // Fallback timeout in case the script hangs or never calls fs.writeFileSync
      setTimeout(() => {
        setIsRunning(prev => {
          if (prev) {
            addLog('error', 'Execution timed out or fs.writeFileSync was never called.');
            return false;
          }
          return prev;
        });
      }, 5000);

    } catch (err: any) {
      addLog('error', `Execution Error: ${err.message || String(err)}`);
      setIsRunning(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#0A0A0A] text-[#F0F0F0] font-sans flex flex-col overflow-hidden">
      {/* Top Nav */}
      <nav className="h-20 border-b border-white/5 px-8 flex items-center justify-between bg-[#0A0A0A] shrink-0">
        <div className="flex items-center gap-4">
          <div className="w-10 h-10 bg-white flex items-center justify-center rounded-sm rotate-3">
            <span className="text-black font-black text-xl">D.</span>
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight uppercase leading-none">DOCX Engine</h1>
            <span className="text-[10px] text-white/40 tracking-[0.2em] uppercase font-mono">v2.4.0 // SANDBOXED</span>
          </div>
        </div>
        <div className="flex items-center gap-8">
          <div className="flex flex-col items-end">
            <span className="text-[10px] text-white/40 uppercase tracking-widest">Execution Environment</span>
            <span className="text-xs font-mono text-emerald-400">V8_NODE_MOCK_READY</span>
          </div>
          <button
            onClick={runScript}
            disabled={isRunning}
            className="flex items-center gap-2 px-8 py-3 bg-[#EFFF00] text-black font-black uppercase text-sm tracking-widest rounded-full hover:scale-105 active:scale-95 transition-transform disabled:opacity-70 disabled:cursor-not-allowed"
          >
            {isRunning ? (
              <Loader2 className="w-4 h-4 animate-spin text-black" />
            ) : null}
            <span>{isRunning ? 'Executing...' : 'Run & Export'}</span>
          </button>
        </div>
      </nav>

      {/* Main Workspace */}
      <main className="flex-1 flex p-6 gap-6 min-h-0">
        
        {/* Editor Column */}
        <section className="flex-1 flex flex-col bg-[#111111] border border-white/10 rounded-xl overflow-hidden relative">
          <div className="h-10 bg-[#161616] border-b border-white/5 flex items-center px-4 justify-between relative z-10 w-full shrink-0">
            <div className="flex gap-2">
              <div className="w-2.5 h-2.5 rounded-full bg-white/20"></div>
              <div className="w-2.5 h-2.5 rounded-full bg-white/20"></div>
              <div className="w-2.5 h-2.5 rounded-full bg-white/20"></div>
            </div>
            <span className="text-[10px] font-mono text-white/30 uppercase">editor.js</span>
            <div className="flex gap-4">
              <button onClick={handleClear} className="text-[10px] text-white/40 hover:text-white uppercase font-bold tracking-tighter transition-colors">Clear</button>
              <button onClick={handlePaste} className="text-[10px] text-white/40 hover:text-white uppercase font-bold tracking-tighter transition-colors">Paste</button>
            </div>
          </div>
          
          <div className="flex-1 overflow-auto bg-transparent relative z-10">
            <Editor
              value={code}
              onValueChange={code => setCode(code)}
              highlight={code => Prism.highlight(code, Prism.languages.javascript, 'javascript')}
              padding={32}
              style={{
                fontFamily: '"JetBrains Mono", "Fira Code", monospace',
                fontSize: 14,
                lineHeight: 1.6,
                minHeight: '100%',
              }}
              textareaClassName="focus:outline-none"
              className="min-h-full text-white/80"
            />
          </div>
          <div className="absolute bottom-8 right-8 w-64 h-64 bg-[#EFFF00] opacity-5 blur-[120px] rounded-full pointer-events-none"></div>
        </section>

        {/* Console Column */}
        <aside className="w-[340px] flex flex-col gap-4 shrink-0">
          <div className="flex-1 bg-black border border-white/10 rounded-xl overflow-hidden flex flex-col">
            <div className="h-10 bg-[#161616] border-b border-white/5 flex items-center px-4 shrink-0">
              <span className="text-[10px] font-mono text-white/60 uppercase tracking-widest">Log Terminal</span>
            </div>
            <div className="flex-1 p-5 font-mono text-[12px] space-y-3 overflow-y-auto">
              <div className="text-white/40"><span className="text-white">System:</span> Runtime initialized.</div>
              <div className="text-white/40"><span className="text-white">System:</span> Buffer polyfill injected.</div>
              <div className="text-white/40"><span className="text-white">System:</span> Intercepting require('docx')...</div>
              
              <div className="pt-4 border-t border-white/5">
                {logs.map((log, idx) => (
                  <div 
                    key={idx} 
                    className={`flex gap-2 break-words mb-2 ${
                      log.type === 'error' ? 'text-red-400' : 
                      log.type === 'success' ? 'text-emerald-400 bg-emerald-400/5 p-2 rounded border border-emerald-400/20' : 
                      'text-white/60'
                    }`}
                  >
                    <div className="shrink-0">
                      {log.type === 'error' ? '$' : 
                       log.type === 'success' ? '' : 
                       <span className="text-blue-400">$</span>}
                    </div>
                    <span className="whitespace-pre-wrap">{log.text}</span>
                  </div>
                ))}
                
                {isRunning && (
                  <div className="flex items-center gap-2 text-emerald-500 animate-pulse mt-2">
                    <span className="animate-pulse text-emerald-300">●</span>
                    <span>Executing...</span>
                  </div>
                )}
              </div>
            </div>
          </div>
          
          <div className="h-32 bg-[#111111] border border-white/10 rounded-xl p-4 flex flex-col justify-between relative overflow-hidden">
            <div className="relative z-10">
              <div className="text-[10px] text-white/40 uppercase tracking-widest mb-1">Memory Usage</div>
              <div className="text-2xl font-bold font-mono tracking-tighter">14.2 <span className="text-sm text-white/40">MB</span></div>
            </div>
            <div className="flex gap-0.5 mt-2">
              <div className="h-1 bg-emerald-400 w-8"></div>
              <div className="h-1 bg-emerald-400 w-10"></div>
              <div className="h-1 bg-emerald-400 w-6"></div>
              <div className="h-1 bg-white/10 w-full"></div>
            </div>
            <div className="absolute -right-4 -bottom-4 text-[80px] font-black text-white/5 italic select-none pointer-events-none">V8</div>
          </div>
        </aside>
        
      </main>
      
      <footer className="h-8 px-6 bg-black border-t border-white/5 flex items-center justify-between text-[10px] text-white/20 uppercase tracking-widest shrink-0 font-mono italic">
        <span>// Execution limit: 5000ms</span>
        <span>Sandboxed Node Context via Browser Polyfill v2.0.4</span>
      </footer>
    </div>
  );
}

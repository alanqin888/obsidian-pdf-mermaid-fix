import { App, MarkdownView, Notice, Plugin, PluginSettingTab, Setting, TFile } from 'obsidian';
import { exec } from 'child_process';
import * as path from 'path';

interface PluginSettings {
	pythonScriptPath: string;
}

const DEFAULT_SETTINGS: PluginSettings = {
	pythonScriptPath: ''
}

export default class PdfMermaidFixPlugin extends Plugin {
	settings: PluginSettings;

	async onload() {
		await this.loadSettings();

		this.addSettingTab(new MdExportSettingTab(this.app, this));

		this.addCommand({
			id: 'export-to-word',
			name: 'Export active file to Word',
			checkCallback: (checking: boolean) => {
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView && markdownView.file) {
					if (!checking) {
						this.exportToWord(markdownView.file);
					}
					return true;
				}
				return false;
			}
		});

		this.addCommand({
			id: 'export-to-pdf-mermaid-fix',
			name: 'Export active file to PDF (Mermaid Fix)',
			checkCallback: (checking: boolean) => {
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView && markdownView.file) {
					if (!checking) {
						this.exportToPdf(markdownView.file);
					}
					return true;
				}
				return false;
			}
		});

		// 注册文件浏览器右键菜单
		this.registerEvent(
			this.app.workspace.on('file-menu', (menu, file) => {
				if (file instanceof TFile && file.extension === 'md') {
					menu.addItem((item) => {
						item.setTitle('Export to Word')
							.setIcon('document')
							.onClick(() => this.exportToWord(file));
					});
					menu.addItem((item) => {
						item.setTitle('Export to PDF (Mermaid Fix)')
							.setIcon('pdf-file')
							.onClick(() => this.exportToPdf(file));
					});
				}
			})
		);

		// 注册编辑器右键菜单
		this.registerEvent(
			this.app.workspace.on('editor-menu', (menu, editor, view) => {
				if (view && view.file) {
					const file = view.file;
					menu.addItem((item) => {
						item.setTitle('Export to Word')
							.setIcon('document')
							.onClick(() => this.exportToWord(file));
					});
					menu.addItem((item) => {
						item.setTitle('Export to PDF (Mermaid Fix)')
							.setIcon('pdf-file')
							.onClick(() => this.exportToPdf(file));
					});
				}
			})
		);
	}

	onunload() {
		// 移除注入的 CSS
		const styleEl = document.getElementById('mermaid-pdf-fix');
		if (styleEl) {
			styleEl.remove();
		}
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	async exportToPdf(file: TFile) {
		new Notice('Applying Mermaid PDF Fix and opening native export...');
		try {
			// 1. Inject the comprehensive print CSS fix for Mermaid & page layout
			let styleEl = document.getElementById('mermaid-pdf-fix');
			if (!styleEl) {
				styleEl = document.createElement('style');
				styleEl.id = 'mermaid-pdf-fix';
				document.head.appendChild(styleEl);
			}
			styleEl.textContent = `
			@media print {
				/* 防止标题及其包裹层孤立在页尾 (Avoid orphan headings) */
				h1, h2, h3, h4, h5, h6,
				.heading-wrapper,
				.markdown-rendered h1,
				.markdown-rendered h2,
				.markdown-rendered h3,
				.markdown-rendered h4,
				.markdown-rendered .heading-wrapper {
					break-after: avoid !important;
					page-break-after: avoid !important;
				}

				/* Mermaid 容器样式优化 */
				.mermaid, .block-language-mermaid, div[data-type="mermaid"] {
					break-inside: avoid !important;
					page-break-inside: avoid !important;
					display: flex !important;
					justify-content: center !important;
					align-items: center !important;
					margin: 0.8em auto !important;
				}

				/* 限制 SVG 宽高：最大高度设为 18cm，确保标题 + 流程图能完美在一页 A4 内放下 */
				.mermaid svg, .block-language-mermaid svg, div[data-type="mermaid"] svg {
					max-width: 100% !important;
					max-height: 18cm !important;
					width: auto !important;
					height: auto !important;
					object-fit: contain !important;
					display: block !important;
					margin: 0 auto !important;
				}

				/* 防止表格与代码块断裂 */
				table, pre {
					break-inside: avoid !important;
					page-break-inside: avoid !important;
				}
			}
			`;

			// 2. We use Obsidian's native PDF export command
			const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
			if (activeView && activeView.file === file) {
				(this.app as any).commands.executeCommandById('workspace:export-pdf');
			} else {
				new Notice('Please open the file to export it to PDF.');
			}
		} catch (error) {
			console.error(error);
			new Notice('Failed to trigger PDF export');
		}
	}

	async exportToWord(file: TFile) {
		new Notice('Starting Word export via Python script...');
		try {
			const adapter = this.app.vault.adapter as any;
			if (!adapter.getBasePath) {
				new Notice('Error: Cannot determine vault absolute path.');
				return;
			}
			const basePath = adapter.getBasePath();
			const absoluteInputPath = path.join(basePath, file.path);
			const absoluteOutputPath = absoluteInputPath.replace(/\.md$/, '.docx');

			// 动态解析自带的脚本路径，如果设置项为空，则默认使用插件目录下的脚本
			const pluginDir = path.join(basePath, this.app.vault.configDir, 'plugins', this.manifest.id);
			const bundledScriptPath = path.join(pluginDir, 'md_to_docx.py');
			const scriptPath = this.settings.pythonScriptPath || bundledScriptPath;
			
			// 解决 macOS GUI 应用环境变量 PATH 缺失的问题，优先尝试 Homebrew 路径，最后尝试系统自带 python3
			const pythonPaths = ['/opt/homebrew/bin/python3', '/usr/local/bin/python3', 'python3'];
			
			const tryRun = (index: number) => {
				if (index >= pythonPaths.length) {
					new Notice('Word export failed. Python3 not found in standard paths.');
					return;
				}
				const py = pythonPaths[index];
				const cmd = `"${py}" "${scriptPath}" "${absoluteInputPath}" -o "${absoluteOutputPath}"`;

				exec(cmd, (error, stdout, stderr) => {
					if (error) {
						console.error(`exec error with ${py}: ${error}`);
						
						// 127 通常表示命令未找到 (Command not found)
						if (error.code === 127 || error.message.includes('not found') || error.message.includes('ENOENT')) {
							tryRun(index + 1);
							return;
						}

						// 检查是否缺少必要的 Python 库
						const fullErr = error.message + '\n' + (stderr || '');
						if (fullErr.includes("No module named 'docx'")) {
							new Notice('Word export failed: Missing Python package "python-docx". Please run "pip install python-docx Pillow" in terminal.', 10000);
						} else if (fullErr.includes("No module named 'PIL'") || fullErr.includes("No module named 'Pillow'")) {
							new Notice('Word export failed: Missing Python package "Pillow". Please run "pip install python-docx Pillow" in terminal.', 10000);
						} else {
							new Notice(`Word export failed: ${error.message}`);
						}
						return;
					}
					if (stderr) {
						console.warn(`stderr: ${stderr}`);
					}
					new Notice('Word export complete! Saved as: ' + file.path.replace(/\.md$/, '.docx'));
				});
			};

			tryRun(0);
		} catch (error) {
			console.error(error);
			new Notice('An error occurred while preparing Word export.');
		}
	}
}

class MdExportSettingTab extends PluginSettingTab {
	plugin: PdfMermaidFixPlugin;

	constructor(app: App, plugin: PdfMermaidFixPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const {containerEl} = this;
		containerEl.empty();

		new Setting(containerEl)
			.setName('Python script path')
			.setDesc('Absolute path to your md_to_docx.py script (Leave empty to use the built-in bundled script).')
			.addText(text => text
				.setPlaceholder('Leave empty to use default bundled script')
				.setValue(this.plugin.settings.pythonScriptPath)
				.onChange(async (value) => {
					this.plugin.settings.pythonScriptPath = value;
					await this.plugin.saveSettings();
				}));
	}
}

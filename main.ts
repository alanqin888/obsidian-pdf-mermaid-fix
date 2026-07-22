import { App, MarkdownView, Notice, Plugin, PluginSettingTab, Setting, TFile } from 'obsidian';
import { exec } from 'child_process';
import * as path from 'path';

interface PluginSettings {
	pythonScriptPath: string;
}

const DEFAULT_SETTINGS: PluginSettings = {
	pythonScriptPath: '/Users/alanqin/.hermes/skills/md-to-docx/scripts/md_to_docx.py'
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
		if (!this.settings.pythonScriptPath) {
			new Notice('Please configure the Python script path in settings first.');
			return;
		}
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

			const scriptPath = this.settings.pythonScriptPath;
			const cmd = `python3 "${scriptPath}" "${absoluteInputPath}" -o "${absoluteOutputPath}"`;

			exec(cmd, (error, stdout, stderr) => {
				if (error) {
					console.error(`exec error: ${error}`);
					new Notice('Word export failed. Check console for details.');
					return;
				}
				if (stderr) {
					console.warn(`stderr: ${stderr}`);
				}
				new Notice('Word export complete! Saved as: ' + file.path.replace(/\.md$/, '.docx'));
			});
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
			.setDesc('Absolute path to your md_to_docx.py script (Required for Word export).')
			.addText(text => text
				.setPlaceholder('/path/to/md_to_docx.py')
				.setValue(this.plugin.settings.pythonScriptPath)
				.onChange(async (value) => {
					this.plugin.settings.pythonScriptPath = value;
					await this.plugin.saveSettings();
				}));
	}
}

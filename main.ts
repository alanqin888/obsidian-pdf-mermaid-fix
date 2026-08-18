import { App, FileSystemAdapter, MarkdownView, Notice, Plugin, PluginSettingTab, Setting, TFile } from 'obsidian';
import { exec } from 'child_process';
import * as path from 'path';

interface PluginSettings {
	pythonScriptPath: string;
}

const DEFAULT_SETTINGS: PluginSettings = {
	pythonScriptPath: ''
}

interface ObsidianApp extends App {
	commands: {
		executeCommandById(id: string): void;
	};
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
						void this.exportToWord(markdownView.file);
					}
					return true;
				}
				return false;
			}
		});

		this.addCommand({
			id: 'export-to-pdf',
			name: 'Export active file to PDF (Mermaid Fix)',
			checkCallback: (checking: boolean) => {
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView && markdownView.file) {
					if (!checking) {
						void this.exportToPdf(markdownView.file);
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
							.onClick(() => {
								void this.exportToWord(file);
							});
					});
					menu.addItem((item) => {
						item.setTitle('Export to PDF (Mermaid Fix)')
							.setIcon('pdf-file')
							.onClick(() => {
								void this.exportToPdf(file);
							});
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
							.onClick(() => {
								void this.exportToWord(file);
							});
					});
					menu.addItem((item) => {
						item.setTitle('Export to PDF (Mermaid Fix)')
							.setIcon('pdf-file')
							.onClick(() => {
								void this.exportToPdf(file);
							});
					});
				}
			})
		);
	}

	onunload() {
		// No style elements to clean up because we use static styles.css
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData()) as PluginSettings;
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	async exportToPdf(file: TFile) {
		new Notice('Applying Mermaid PDF Fix and generating TOC for export...');
		try {
			const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
			if (!activeView || activeView.file !== file) {
				new Notice('Please open the file to export it to PDF.');
				return;
			}

			// 1. 扫描当前视图中的所有标题并赋予唯一 ID 锚点
			const container = activeView.contentEl;
			const headings = Array.from(container.querySelectorAll<HTMLElement>('h1, h2, h3, .markdown-rendered h1, .markdown-rendered h2, .markdown-rendered h3'));
			
			let tocContainer: HTMLElement | null = null;

			if (headings.length > 0) {
				tocContainer = document.createElement('div');
				tocContainer.className = 'pdf-toc-container';

				const titleEl = document.createElement('div');
				titleEl.className = 'pdf-toc-title';
				titleEl.textContent = '目录 / Table of Contents';
				tocContainer.appendChild(titleEl);

				const listEl = document.createElement('ul');
				listEl.className = 'pdf-toc-list';

				headings.forEach((heading, idx) => {
					// 忽略隐藏或原本就在目录容器里的标题
					if (heading.closest('.pdf-toc-container')) return;

					let headingId = heading.id;
					if (!headingId) {
						headingId = `pdf-heading-${idx}-${Date.now()}`;
						heading.id = headingId;
					}

					const tag = heading.tagName.toLowerCase();
					const itemEl = document.createElement('li');
					itemEl.className = `pdf-toc-item pdf-toc-item-${tag}`;

					const linkEl = document.createElement('a');
					linkEl.className = 'pdf-toc-link';
					linkEl.href = `#${headingId}`;
					linkEl.textContent = heading.textContent ? heading.textContent.trim() : '';

					itemEl.appendChild(linkEl);
					listEl.appendChild(itemEl);
				});

				tocContainer.appendChild(listEl);

				// 插入到预览容器的最上方
				const previewEl = container.querySelector('.markdown-preview-view') || container;
				previewEl.prepend(tocContainer);
			}

			// 2. 调用 Obsidian 原生 PDF 导出命令
			(this.app as ObsidianApp).commands.executeCommandById('workspace:export-pdf');

			// 3. 导出触发后清理临时目录 DOM
			setTimeout(() => {
				if (tocContainer && tocContainer.parentNode) {
					tocContainer.parentNode.removeChild(tocContainer);
				}
			}, 2000);
		} catch (error) {
			const err = error as Error;
			console.error(err);
			new Notice(`Failed to trigger PDF export: ${err.message}`);
		}
	}

	async exportToWord(file: TFile) {
		new Notice('Starting Word export via Python script...');
		try {
			const adapter = this.app.vault.adapter;
			if (!(adapter instanceof FileSystemAdapter)) {
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
						console.error(`exec error with ${py}: ${error.message}`);
						
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
			const err = error as Error;
			console.error(err);
			new Notice(`An error occurred while preparing Word export: ${err.message}`);
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

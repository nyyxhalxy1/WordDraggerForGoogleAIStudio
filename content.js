/**
 * DOCX to Google AI Studio — Content Script (Isolated World)
 *
 * 在 Google AI Studio 页面中拦截 .docx 文件的拖放和文件选择操作，
 * 将其转换为 Markdown（+ 可选图片）后触发原生上传。
 */

(function () {
  'use strict';

  // =========================================================================
  // 常量 & 工具
  // =========================================================================

  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const DOCX_EXT = '.docx';

  const MIME_MAP = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tiff: 'image/tiff',
    tif: 'image/tiff',
    svg: 'image/svg+xml',
    webp: 'image/webp',
  };

  /** 判断文件是否为 .docx */
  function isDocx(file) {
    if (!file) return false;
    if (file.name && file.name.toLowerCase().endsWith(DOCX_EXT)) return true;
    if (file.type === DOCX_MIME) return true;
    return false;
  }

  /** 根据扩展名获取 MIME 类型 */
  function getMimeType(ext) {
    return MIME_MAP[(ext || '').toLowerCase()] || 'application/octet-stream';
  }

  /** 日志辅助 */
  function log(...args) {
    console.log('[DOCX→AI Studio]', ...args);
  }

  function logError(...args) {
    console.error('[DOCX→AI Studio]', ...args);
  }

  // =========================================================================
  // UI: 加载指示器
  // =========================================================================

  function showLoading(fileName) {
    removeLoading();
    const el = document.createElement('div');
    el.id = 'docx-converter-loading';
    el.innerHTML = `
      <div class="docx-loading-overlay">
        <div class="docx-loading-content">
          <div class="docx-loading-spinner"></div>
          <p>正在转换: <strong>${escapeHtml(fileName)}</strong></p>
          <p class="docx-loading-sub">正在将 Word 文档转换为 Markdown ...</p>
        </div>
      </div>`;
    document.body.appendChild(el);
  }

  function removeLoading() {
    const el = document.getElementById('docx-converter-loading');
    if (el) el.remove();
  }

  // =========================================================================
  // UI: 图片选择对话框
  // =========================================================================

  function showImageDialog(fileName, markdown, images) {
    return new Promise((resolve) => {
      removeDialog();
      const el = document.createElement('div');
      el.id = 'docx-converter-dialog';
      const mdSizeKB = (new Blob([markdown]).size / 1024).toFixed(1);
      el.innerHTML = `
        <div class="docx-dialog-overlay">
          <div class="docx-dialog-content">
            <div class="docx-dialog-header">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" style="flex-shrink:0">
                <path d="M14 2H6C4.9 2 4 2.9 4 4V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V8L14 2Z" fill="#4285F4"/>
                <path d="M14 2V8H20" fill="#A1C2FA"/>
                <path d="M16 13H8V15H16V13ZM16 17H8V19H16V17ZM14 9H8V11H14V9Z" fill="white"/>
              </svg>
              <h3>${escapeHtml(fileName)}</h3>
            </div>
            <div class="docx-dialog-body">
              <p>文档包含 <strong>${images.length}</strong> 张图片</p>
              <p class="docx-dialog-info">Markdown 文本大小: ${mdSizeKB} KB</p>
            </div>
            <div class="docx-dialog-buttons">
              <button id="docx-btn-text-only" class="docx-btn docx-btn-secondary">
                <span class="docx-btn-icon">📄</span>
                <span class="docx-btn-label">
                  <strong>仅上传文本</strong>
                  <small>丢弃图片，只上传 Markdown（1 个文件）</small>
                </span>
              </button>
              <button id="docx-btn-text-images" class="docx-btn docx-btn-primary">
                <span class="docx-btn-icon">📦</span>
                <span class="docx-btn-label">
                  <strong>上传文本 + 图片</strong>
                  <small>上传 Markdown + ${images.length} 张图片（共 ${images.length + 1} 个文件）</small>
                </span>
              </button>
              <button id="docx-btn-cancel" class="docx-btn docx-btn-ghost">取消</button>
            </div>
          </div>
        </div>`;
      document.body.appendChild(el);

      // 事件绑定
      const cleanup = (result) => {
        el.remove();
        resolve(result);
      };

      document.getElementById('docx-btn-text-only').addEventListener('click', () => cleanup('text-only'));
      document.getElementById('docx-btn-text-images').addEventListener('click', () => cleanup('text-images'));
      document.getElementById('docx-btn-cancel').addEventListener('click', () => cleanup('cancel'));

      // ESC 键关闭
      const onKeyDown = (e) => {
        if (e.key === 'Escape') {
          document.removeEventListener('keydown', onKeyDown);
          cleanup('cancel');
        }
      };
      document.addEventListener('keydown', onKeyDown);

      // 点击遮罩关闭
      el.querySelector('.docx-dialog-overlay').addEventListener('click', (e) => {
        if (e.target === e.currentTarget) cleanup('cancel');
      });
    });
  }

  function removeDialog() {
    const el = document.getElementById('docx-converter-dialog');
    if (el) el.remove();
  }

  // =========================================================================
  // UI: 通知
  // =========================================================================

  function showNotification(message, type = 'success') {
    const el = document.createElement('div');
    el.className = `docx-notification docx-notification-${type}`;
    el.textContent = message;
    document.body.appendChild(el);
    // 触发进入动画
    requestAnimationFrame(() => el.classList.add('docx-notification-show'));
    setTimeout(() => {
      el.classList.remove('docx-notification-show');
      setTimeout(() => el.remove(), 300);
    }, 3000);
  }

  // =========================================================================
  // 核心: 解析 .docx → Markdown + 图片
  // =========================================================================

  /**
   * 解析 .docx 文件，返回 { markdown: string, images: Array<{name, blob, mimeType}> }
   */
  async function parseDocx(file) {
    const arrayBuffer = await file.arrayBuffer();

    // 并行执行: mammoth 转 HTML + JSZip 提取图片
    const [htmlResult, images] = await Promise.all([
      convertToHtml(arrayBuffer),
      extractImages(arrayBuffer),
    ]);

    // HTML → Markdown
    const markdown = htmlToMarkdown(htmlResult.html, file.name);

    log(`转换完成: ${file.name}`, {
      markdownLength: markdown.length,
      imageCount: images.length,
      warnings: htmlResult.warnings,
    });

    return { markdown, images };
  }

  /** 用 mammoth 将 docx ArrayBuffer 转为 HTML */
  async function convertToHtml(arrayBuffer) {
    const warnings = [];
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        // 图片处理：不内联 base64，只保留引用
        convertImage: mammoth.images.imgElement(function (image) {
          return image.read('base64').then(function (imageBuffer) {
            // 用占位符标记，后续在 markdown 中替换
            const ext = (image.contentType || 'image/png').split('/')[1] || 'png';
            return {
              src: `__DOCX_IMG_placeholder__.${ext}`,
              alt: `image.${ext}`,
            };
          });
        }),
        styleMap: [
          // 增强样式映射
          "p[style-name='Title'] => h1:fresh",
          "p[style-name='Subtitle'] => h2:fresh",
          "p[style-name='Heading 1'] => h1:fresh",
          "p[style-name='Heading 2'] => h2:fresh",
          "p[style-name='Heading 3'] => h3:fresh",
          "p[style-name='Heading 4'] => h4:fresh",
          "p[style-name='Heading 5'] => h5:fresh",
          "p[style-name='Heading 6'] => h6:fresh",
        ],
      }
    );
    if (result.messages) {
      result.messages.forEach((m) => {
        if (m.type === 'warning') warnings.push(m.message);
      });
    }
    return { html: result.value, warnings };
  }

  /** 用 JSZip 提取 docx 中的图片文件 */
  async function extractImages(arrayBuffer) {
    const images = [];
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const mediaFiles = [];

      zip.forEach((relativePath, zipEntry) => {
        if (
          !zipEntry.dir &&
          relativePath.startsWith('word/media/') &&
          relativePath !== 'word/media/'
        ) {
          mediaFiles.push({ path: relativePath, entry: zipEntry });
        }
      });

      for (const { path, entry } of mediaFiles) {
        try {
          const blob = await entry.async('blob');
          const fileName = path.split('/').pop();
          const ext = fileName.split('.').pop().toLowerCase();
          const mimeType = getMimeType(ext);

          // 跳过浏览器不支持的格式
          if (ext === 'emf' || ext === 'wmf') {
            log(`跳过不支持的图片格式: ${fileName}`);
            continue;
          }

          images.push({ name: fileName, blob, mimeType });
        } catch (err) {
          logError(`提取图片失败: ${path}`, err);
        }
      }
    } catch (err) {
      logError('JSZip 解析失败:', err);
    }
    return images;
  }

  /** 将 mammoth 输出的 HTML 转为 Markdown */
  function htmlToMarkdown(html, fileName) {
    // 配置 Turndown
    const turndownService = new TurndownService({
      headingStyle: 'atx',
      codeBlockStyle: 'fenced',
      bulletListMarker: '-',
      emDelimiter: '*',
      strongDelimiter: '**',
      hr: '---',
    });

    // 启用 GFM 表格和删除线
    if (typeof turndownPluginGfm !== 'undefined') {
      const gfm = turndownPluginGfm.gfm;
      turndownService.use(gfm);
    }

    // 自定义规则：处理图片占位符 img 标签
    turndownService.addRule('docxImagePlaceholder', {
      filter: function (node) {
        return (
          node.nodeName === 'IMG' &&
          node.getAttribute('src') &&
          node.getAttribute('src').includes('__DOCX_IMG_placeholder__')
        );
      },
      replacement: function (content, node) {
        const alt = node.getAttribute('alt') || 'image';
        const src = node.getAttribute('src');
        // 提取扩展名作为文件名提示
        return `![${alt}](${src})`;
      },
    });

    let markdown = turndownService.turndown(html);

    // 后处理：清理多余空行
    markdown = markdown
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    // 添加文件来源注释头
    const baseName = fileName.replace(/\.docx$/i, '');
    markdown = `<!-- 由 "${fileName}" 转换生成 -->\n\n# ${baseName}\n\n${markdown}`;

    return markdown;
  }

  // =========================================================================
  // 核心: 处理单个 .docx 文件的完整流程
  // =========================================================================

  async function processDocx(docxFile) {
    log(`开始处理: ${docxFile.name} (${(docxFile.size / 1024).toFixed(1)} KB)`);
    showLoading(docxFile.name);

    try {
      const { markdown, images } = await parseDocx(docxFile);
      removeLoading();

      const baseName = docxFile.name.replace(/\.docx$/i, '');
      let filesToUpload = [];

      if (images.length === 0) {
        // 无图片，直接上传 Markdown
        const cleanMarkdown = cleanImagePlaceholders(markdown);
        const mdFile = createMarkdownFile(cleanMarkdown, `${baseName}.md`);
        filesToUpload = [mdFile];
        log('无图片，直接上传 Markdown');
      } else {
        // 有图片，弹出选择对话框
        const choice = await showImageDialog(docxFile.name, markdown, images);

        if (choice === 'cancel') {
          log('用户取消操作');
          showNotification('已取消转换', 'info');
          return;
        }

        if (choice === 'text-only') {
          // 仅文本
          const cleanMarkdown = cleanImagePlaceholders(markdown);
          const mdFile = createMarkdownFile(cleanMarkdown, `${baseName}.md`);
          filesToUpload = [mdFile];
          log('用户选择仅上传文本');
        } else {
          // 文本 + 图片
          // 在 Markdown 中将占位符替换为实际图片文件名
          let mdWithRefs = markdown;
          images.forEach((img, i) => {
            // 替换第 i+1 个占位符为实际文件名
            mdWithRefs = mdWithRefs.replace(
              /!\[([^\]]*)\]\(__DOCX_IMG_placeholder__\.[a-z]+\)/,
              `![${img.name}](${img.name})`
            );
          });
          const mdFile = createMarkdownFile(mdWithRefs, `${baseName}.md`);
          const imageFiles = images.map(
            (img) => new File([img.blob], img.name, { type: img.mimeType })
          );
          filesToUpload = [mdFile, ...imageFiles];
          log(`用户选择上传文本 + ${images.length} 张图片`);
        }
      }

      // 触发上传
      triggerUpload(filesToUpload);
      showNotification(
        `已转换 "${docxFile.name}" → ${filesToUpload.length} 个文件`,
        'success'
      );
    } catch (err) {
      removeLoading();
      logError('处理失败:', err);
      showNotification(`转换失败: ${err.message}`, 'error');
    }
  }

  /** 清理 Markdown 中的图片占位符 */
  function cleanImagePlaceholders(markdown) {
    return markdown.replace(/!\[[^\]]*\]\(__DOCX_IMG_placeholder__\.[a-z]+\)\n*/g, '');
  }

  /** 创建 Markdown File 对象 */
  function createMarkdownFile(content, fileName) {
    const blob = new Blob([content], { type: 'text/markdown; charset=utf-8' });
    return new File([blob], fileName, {
      type: 'text/markdown',
      lastModified: Date.now(),
    });
  }

  // =========================================================================
  // 核心: 触发文件上传到 Google AI Studio
  // =========================================================================

  /**
   * 通过消息通道发送到主世界脚本 (inject.js) 来触发上传
   */
  function triggerUpload(files) {
    log(`准备上传 ${files.length} 个文件:`, files.map((f) => f.name));

    // 将 File 对象序列化为可传递到主世界的格式
    const serializeFiles = files.map((file) => {
      return file.arrayBuffer().then((buf) => ({
        name: file.name,
        type: file.type,
        lastModified: file.lastModified,
        buffer: Array.from(new Uint8Array(buf)), // 转为普通数组以便 JSON 序列化
      }));
    });

    Promise.all(serializeFiles).then((serialized) => {
      // 通过 window.postMessage 发送到主世界
      window.postMessage(
        {
          type: 'DOCX_CONVERTER_UPLOAD',
          files: serialized,
        },
        '*'
      );
      log('已发送上传请求到主世界');
    });
  }

  // =========================================================================
  // 事件拦截: 拖放 (Drag & Drop)
  // =========================================================================

  /** 检查 DataTransfer 中是否包含 .docx 文件 */
  function hasDocxInDataTransfer(dataTransfer) {
    if (!dataTransfer) return false;

    // 检查 items (dragover 阶段)
    if (dataTransfer.items) {
      for (let i = 0; i < dataTransfer.items.length; i++) {
        const item = dataTransfer.items[i];
        if (item.kind === 'file') {
          // dragover 阶段可能无法获取文件名，只能检查 type
          if (item.type === DOCX_MIME) return true;
          // 也尝试通过 webkitGetAsEntry 获取文件名
          try {
            const entry = item.webkitGetAsEntry && item.webkitGetAsEntry();
            if (entry && entry.name && entry.name.toLowerCase().endsWith(DOCX_EXT)) return true;
          } catch (e) { /* 忽略 */ }
        }
      }
    }

    // 检查 files (drop 阶段)
    if (dataTransfer.files) {
      for (let i = 0; i < dataTransfer.files.length; i++) {
        if (isDocx(dataTransfer.files[i])) return true;
      }
    }

    return false;
  }

  // 标记，避免重入
  let isProcessing = false;
  // 标记已转换的 drop 事件，避免循环拦截
  const CONVERTED_FLAG = '__docx_converter_handled__';

  function handleDragOver(e) {
    if (isProcessing) return;
    if (!hasDocxInDataTransfer(e.dataTransfer)) return;

    e.preventDefault();
    e.stopImmediatePropagation();
    e.dataTransfer.dropEffect = 'copy';

    // 显示拖放视觉提示
    showDropHint();
  }

  function handleDragLeave(e) {
    hideDropHint();
  }

  function handleDragEnter(e) {
    if (isProcessing) return;
    if (!hasDocxInDataTransfer(e.dataTransfer)) return;
    e.preventDefault();
    e.stopImmediatePropagation();
  }

  async function handleDrop(e) {
    hideDropHint();

    // 如果是转换后重新派发的事件，不再拦截
    if (e[CONVERTED_FLAG]) return;

    if (isProcessing) return;

    const files = e.dataTransfer?.files;
    if (!files || files.length === 0) return;

    const docxFiles = Array.from(files).filter(isDocx);
    const otherFiles = Array.from(files).filter((f) => !isDocx(f));

    if (docxFiles.length === 0) return; // 无 .docx，让页面正常处理

    // 阻止事件传播给 AI Studio
    e.preventDefault();
    e.stopImmediatePropagation();

    isProcessing = true;

    try {
      // 如果有非 docx 文件，先让它们正常上传
      if (otherFiles.length > 0) {
        passNonDocxFiles(otherFiles, e.target);
      }

      // 依次处理每个 docx 文件
      for (const docxFile of docxFiles) {
        await processDocx(docxFile);
      }
    } finally {
      isProcessing = false;
    }
  }

  /** 将非 docx 文件重新派发给页面处理 */
  function passNonDocxFiles(files, target) {
    try {
      const dt = new DataTransfer();
      files.forEach((f) => dt.items.add(f));
      const newEvent = new DragEvent('drop', {
        bubbles: true,
        cancelable: true,
        dataTransfer: dt,
      });
      newEvent[CONVERTED_FLAG] = true;
      (target || document.body).dispatchEvent(newEvent);
    } catch (err) {
      logError('重新派发非 docx 文件失败:', err);
    }
  }

  // 拖放视觉提示
  let dropHintEl = null;
  function showDropHint() {
    if (dropHintEl) return;
    dropHintEl = document.createElement('div');
    dropHintEl.id = 'docx-drop-hint';
    dropHintEl.innerHTML = `
      <div class="docx-drop-hint-inner">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
          <path d="M14 2H6C4.9 2 4 2.9 4 4V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V8L14 2Z" fill="#4285F4" opacity="0.8"/>
          <path d="M14 2V8H20" fill="#A1C2FA" opacity="0.8"/>
          <path d="M12 18L12 12M12 12L9 15M12 12L15 15" stroke="white" stroke-width="2" stroke-linecap="round"/>
        </svg>
        <p>松开以转换 Word 文档</p>
      </div>`;
    document.body.appendChild(dropHintEl);
  }

  function hideDropHint() {
    if (dropHintEl) {
      dropHintEl.remove();
      dropHintEl = null;
    }
  }

  // =========================================================================
  // 事件拦截: 文件选择器 (<input type="file">)
  // =========================================================================

  /** 监听 file input 的 change 事件 */
  function interceptFileInput(input) {
    if (input._docxIntercepted) return;
    input._docxIntercepted = true;

    input.addEventListener(
      'change',
      async function (e) {
        if (isProcessing) return;
        if (!input.files || input.files.length === 0) return;

        const docxFiles = Array.from(input.files).filter(isDocx);
        if (docxFiles.length === 0) return; // 无 .docx，不拦截

        const otherFiles = Array.from(input.files).filter((f) => !isDocx(f));

        // 阻止原始事件
        e.preventDefault();
        e.stopImmediatePropagation();

        isProcessing = true;
        try {
          // 处理 docx 文件
          for (const docxFile of docxFiles) {
            await processDocx(docxFile);
          }

          // 如果有非 docx 文件，设置回 input 并触发 change
          if (otherFiles.length > 0) {
            const dt = new DataTransfer();
            otherFiles.forEach((f) => dt.items.add(f));
            input.files = dt.files;
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }
        } finally {
          isProcessing = false;
        }
      },
      true // 捕获阶段
    );
  }

  /** MutationObserver 监视新增的 file input */
  function observeFileInputs() {
    // 拦截已存在的
    document.querySelectorAll('input[type="file"]').forEach(interceptFileInput);

    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of mutation.addedNodes) {
          if (node.nodeType !== Node.ELEMENT_NODE) continue;
          if (node.matches && node.matches('input[type="file"]')) {
            interceptFileInput(node);
          }
          // 也检查子元素
          if (node.querySelectorAll) {
            node.querySelectorAll('input[type="file"]').forEach(interceptFileInput);
          }
        }
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
    log('MutationObserver 已启动，监控 file input');
  }

  // =========================================================================
  // HTML 转义
  // =========================================================================

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // =========================================================================
  // 初始化
  // =========================================================================

  function init() {
    log('扩展已加载，URL:', location.href);

    // 注册拖放拦截（捕获阶段）
    document.addEventListener('dragover', handleDragOver, true);
    document.addEventListener('dragenter', handleDragEnter, true);
    document.addEventListener('dragleave', handleDragLeave, true);
    document.addEventListener('drop', handleDrop, true);

    // 观察 file input
    observeFileInputs();

    log('拖放拦截和文件选择器监控已就绪');
  }

  // 确保 DOM 就绪后初始化
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();

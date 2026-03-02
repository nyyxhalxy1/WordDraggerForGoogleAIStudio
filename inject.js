/**
 * DOCX to Google AI Studio — Main World Script (inject.js)
 *
 * 此脚本注入到页面的主世界（MAIN world），可以直接访问页面的 DOM 和 JS 上下文。
 * 职责：接收来自内容脚本的消息，构造 File 对象，并触发 Google AI Studio 的文件上传。
 */

(function () {
  'use strict';

  const LOG_PREFIX = '[DOCX→AI Studio|MAIN]';

  function log(...args) {
    console.log(LOG_PREFIX, ...args);
  }

  function logError(...args) {
    console.error(LOG_PREFIX, ...args);
  }

  // =========================================================================
  // 文件上传触发
  // =========================================================================

  /**
   * 查找 Google AI Studio 的文件上传目标区域
   * 通过多种策略依次尝试定位
   */
  function findDropTarget() {
    const selectors = [
      // AI Studio 常见的拖放目标选择器（按可能性排序）
      // 聊天输入区域
      'ms-autosize-textarea',
      'ms-chunk-editor',
      '.prompt-input-container',
      '.text-input-field',
      'textarea[aria-label]',
      // 对话区域
      '.conversation-container',
      'ms-prompt-chunk',
      '.chat-container',
      // Angular Material 容器
      'mat-sidenav-content',
      '.mat-sidenav-content',
      // 通用容器
      'main',
      '[role="main"]',
      '.main-content',
    ];

    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) {
        log('找到上传目标:', sel);
        return el;
      }
    }

    log('未找到特定上传目标，使用 document.body');
    return document.body;
  }

  /**
   * 查找页面中的 <input type="file"> 元素
   */
  function findFileInput() {
    // 优先查找可见或最近添加的 file input
    const inputs = document.querySelectorAll('input[type="file"]');
    if (inputs.length > 0) {
      // 返回最后一个（通常是最新创建的）
      return inputs[inputs.length - 1];
    }
    return null;
  }

  /**
   * 策略 A：通过派发 drop 事件序列触发上传
   */
  function uploadViaDrop(files) {
    const target = findDropTarget();

    try {
      const dataTransfer = new DataTransfer();
      files.forEach((f) => dataTransfer.items.add(f));

      // 模拟完整的拖放事件序列
      const commonOpts = {
        bubbles: true,
        cancelable: true,
        dataTransfer: dataTransfer,
      };

      target.dispatchEvent(new DragEvent('dragenter', commonOpts));
      target.dispatchEvent(new DragEvent('dragover', commonOpts));

      const dropEvent = new DragEvent('drop', commonOpts);
      target.dispatchEvent(dropEvent);

      target.dispatchEvent(new DragEvent('dragleave', commonOpts));

      log('策略 A (Drop 事件) 已执行');
      return true;
    } catch (err) {
      logError('策略 A 失败:', err);
      return false;
    }
  }

  /**
   * 策略 B：通过设置 <input type="file">.files 触发上传
   */
  function uploadViaFileInput(files) {
    const fileInput = findFileInput();
    if (!fileInput) {
      log('策略 B: 未找到 file input');
      return false;
    }

    try {
      const dataTransfer = new DataTransfer();
      files.forEach((f) => dataTransfer.items.add(f));
      fileInput.files = dataTransfer.files;

      // 触发 change 事件
      fileInput.dispatchEvent(new Event('change', { bubbles: true }));
      // 有些框架监听 input 事件
      fileInput.dispatchEvent(new Event('input', { bubbles: true }));

      log('策略 B (File Input) 已执行');
      return true;
    } catch (err) {
      logError('策略 B 失败:', err);
      return false;
    }
  }

  /**
   * 策略 C：尝试点击上传按钮来触发 file input，然后设置文件
   */
  function uploadViaButtonClick(files) {
    // 尝试找到 "添加文件" 按钮
    const buttons = document.querySelectorAll('button');
    let addFileBtn = null;

    for (const btn of buttons) {
      const text = (btn.textContent || '').toLowerCase();
      const ariaLabel = (btn.getAttribute('aria-label') || '').toLowerCase();
      if (
        text.includes('add file') ||
        text.includes('upload') ||
        text.includes('attach') ||
        ariaLabel.includes('add file') ||
        ariaLabel.includes('upload') ||
        ariaLabel.includes('attach')
      ) {
        addFileBtn = btn;
        break;
      }
    }

    if (!addFileBtn) {
      // 尝试通过图标按钮（常见的 attachment/upload icon）
      const iconBtns = document.querySelectorAll(
        'button mat-icon, button .material-icons, button svg'
      );
      for (const icon of iconBtns) {
        const text = icon.textContent || '';
        if (text === 'attach_file' || text === 'upload_file' || text === 'add') {
          addFileBtn = icon.closest('button');
          break;
        }
      }
    }

    if (!addFileBtn) {
      log('策略 C: 未找到上传按钮');
      return false;
    }

    try {
      // 临时拦截 file input 的打开，设置我们的文件
      const originalClick = HTMLInputElement.prototype.click;
      let intercepted = false;

      HTMLInputElement.prototype.click = function () {
        if (this.type === 'file' && !intercepted) {
          intercepted = true;
          const dt = new DataTransfer();
          files.forEach((f) => dt.items.add(f));
          this.files = dt.files;
          this.dispatchEvent(new Event('change', { bubbles: true }));
          HTMLInputElement.prototype.click = originalClick;
          log('策略 C: 拦截了 file input click');
          return;
        }
        return originalClick.call(this);
      };

      addFileBtn.click();

      // 恢复原始 click（以防回调没有被触发）
      setTimeout(() => {
        HTMLInputElement.prototype.click = originalClick;
      }, 1000);

      log('策略 C (按钮点击) 已执行');
      return true;
    } catch (err) {
      logError('策略 C 失败:', err);
      return false;
    }
  }

  /**
   * 使用多种策略依次尝试触发上传
   */
  function performUpload(files) {
    log(`开始上传 ${files.length} 个文件:`, files.map((f) => `${f.name} (${f.type})`));

    // 策略 A: Drop 事件
    if (uploadViaDrop(files)) return;

    // 策略 B: File Input
    if (uploadViaFileInput(files)) return;

    // 策略 C: 按钮点击
    if (uploadViaButtonClick(files)) return;

    logError('所有上传策略均失败');
  }

  // =========================================================================
  // 消息监听：接收来自隔离世界 content.js 的消息
  // =========================================================================

  window.addEventListener('message', function (event) {
    if (event.source !== window) return;
    if (!event.data || event.data.type !== 'DOCX_CONVERTER_UPLOAD') return;

    log('收到上传请求:', event.data.files?.length, '个文件');

    const serializedFiles = event.data.files;
    if (!serializedFiles || serializedFiles.length === 0) return;

    // 将序列化的数据还原为 File 对象
    const files = serializedFiles.map((sf) => {
      const uint8Array = new Uint8Array(sf.buffer);
      const blob = new Blob([uint8Array], { type: sf.type });
      return new File([blob], sf.name, {
        type: sf.type,
        lastModified: sf.lastModified,
      });
    });

    // 延迟一帧执行上传，确保 DOM 稳定
    requestAnimationFrame(() => {
      performUpload(files);
    });
  });

  log('主世界脚本已加载');
})();

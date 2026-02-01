/**
 * 系統狀態管理器
 * 負責控制 Header 的狀態指示器 (System Ready Indicator)
 */
class StatusManager {
    constructor() {
        this.dom = {
            container: document.getElementById('status-indicator'),
            text: document.getElementById('status-text'),
            dot1: document.getElementById('status-dot-1'),
            dot2: document.getElementById('status-dot-2')
        };

        // 預定義狀態樣式
        this.states = {
            ready: {
                text: '系統就緒',
                dot1: 'bg-primary',
                dot2: 'bg-emerald-500',
                container: 'bg-slate-100 dark:bg-slate-900 border-slate-200 dark:border-white/10',
                animate: false
            },
            processing: {
                text: '處理中...',
                dot1: 'bg-amber-400',
                dot2: 'bg-amber-500',
                container: 'bg-amber-50 dark:bg-amber-900/20 border-amber-200 dark:border-amber-500/30',
                animate: true
            },
            success: {
                text: '處理完成',
                dot1: 'bg-emerald-400',
                dot2: 'bg-emerald-600',
                container: 'bg-emerald-50 dark:bg-emerald-900/20 border-emerald-200 dark:border-emerald-500/30',
                animate: false
            },
            error: {
                text: '發生異常',
                dot1: 'bg-red-500',
                dot2: 'bg-red-600',
                container: 'bg-red-50 dark:bg-red-900/20 border-red-200 dark:border-red-500/30',
                animate: true // Error 也可以有些微動畫
            }
        };

        this.currentAnimInterval = null;
    }

    /**
     * 設置狀態
     * @param {string} stateKey - 'ready' | 'processing' | 'success' | 'error'
     */
    setStatus(stateKey) {
        const state = this.states[stateKey];
        if (!state || !this.dom.container) return;

        // 停止之前的動畫
        this.stopAnimation();

        // 更新文字
        this.dom.text.textContent = state.text;

        // 更新容器樣式 (先移除舊有衝突的 class, 這裡簡化處理直接覆蓋關鍵背景色)
        // 為了乾淨，我們移除所有顏色相關的 class 再加上新的
        this._resetClasses();

        // 應用新樣式
        this._applyClasses(this.dom.container, state.container);
        this._applyClasses(this.dom.dot1, state.dot1);
        this._applyClasses(this.dom.dot2, state.dot2);

        // 如果需要動畫
        if (state.animate) {
            this.startPulseAnimation();
        }

        // 如果是成功或錯誤，設定計時器自動切回 Ready (選擇性)
        if (stateKey === 'success') {
            setTimeout(() => this.setStatus('ready'), 3000);
        }
    }

    _resetClasses() {
        // 清除顏色相關的 Class (保留 layout 相關的)
        // 簡單起見，我們針對已知的顏色 class 進行替換
        const colorClasses = [
            'bg-primary', 'bg-emerald-500', 'bg-slate-100', 'dark:bg-slate-900',
            'bg-amber-400', 'bg-amber-500', 'bg-amber-50', 'dark:bg-amber-900/20',
            'bg-emerald-400', 'bg-emerald-600', 'bg-emerald-50', 'dark:bg-emerald-900/20',
            'bg-red-500', 'bg-red-600', 'bg-red-50', 'dark:bg-red-900/20',
            'border-slate-200', 'dark:border-white/10',
            'border-amber-200', 'dark:border-amber-500/30',
            'border-emerald-200', 'dark:border-emerald-500/30',
            'border-red-200', 'dark:border-red-500/30'
        ];

        [this.dom.container, this.dom.dot1, this.dom.dot2].forEach(el => {
            el.classList.remove(...colorClasses);
        });
    }

    _applyClasses(element, classString) {
        const classes = classString.split(' ');
        element.classList.add(...classes);
    }

    startPulseAnimation() {
        // 簡單的 CSS class 脈衝動畫
        this.dom.container.classList.add('animate-pulse');
    }

    stopAnimation() {
        this.dom.container.classList.remove('animate-pulse');
    }
}

// 全域實例
const statusManager = new StatusManager();

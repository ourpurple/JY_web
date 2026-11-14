// ==UserScript==
// @name         菁优网一键提取试卷（净化页面 + Word/PDF导出）
// @namespace    https://etshelper.ccwu.cc
// @version      2.6.5
// @description  两阶段净化与 Word/PDF 导出；保留题目、段落与强调文字，题干填空留白，支持原生 Word 公式、等比例图片与资源就绪检查。
// @author       https://etshelper.ccwu.cc
// @license      MIT
// @match        *://*.jyeoo.com/*/report/detail/*
// @match        *://*.jyeoo.com/*/report/detail*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function() {
    'use strict';

    // ================= 1. 解除右键与文本复制锁 =================
    function unlockPage() {
        document.body.oncontextmenu = null;
        document.body.oncopy = null;
        document.body.onselectstart = null;
        document.oncontextmenu = null;
        document.onselectstart = null;

        const unlockStyle = document.createElement('style');
        unlockStyle.id = 'jyeoo-unlock-style';
        unlockStyle.innerHTML = `
            * {
                user-select: auto !important;
                -webkit-user-select: auto !important;
            }
            .sh { display: block !important; }
        `;
        document.head.appendChild(unlockStyle);
    }
    unlockPage();

    // ================= 2. 试卷区域与内容保真 =================
    const PAPER_SELECTORS = ['#divReportMain', '.report-main', '.new-content-left', '.report-main-box', '.paper-detail'];
    const QUESTION_SELECTOR = 'fieldset, .quesborder, .QUES_LI, .ques-detail, .ques-item, .question-item';
    const STEM_SELECTOR = '.pt1, .ques-con, .q-item';
    const MATH_SELECTOR = '.MathJye, [class*="MathJye"]';
    const OPTION_SELECTOR = 'table.composition, table.composition2';
    const ANSWER_SELECTOR = '.sanwei-con, .sanwei-box, .pt-answer, .pt-analysis, .pt-detail, .answer-box, .detail-con, .pt-points, .pt-degree, .pt-source, .pt-knowledges, .ques-report, .detail-item, .ans-details';
    // PDF 专用：按用户实测 CSS 10pt → PDF 7.44pt 校准；Word 字号独立设置。
    const PDF_TYPOGRAPHY = Object.freeze({ bodyPt: 9, titlePt: 10, measuredScale: 7.44 / 10 });
    const SITE_BLANK_SELECTOR = '.quizPutTag, .quiz-put-tag';
    const EXPLICIT_BLANK_SELECTOR = `${SITE_BLANK_SELECTOR}, .mathjye-underline, .fill-blank, .blank, .tkt, input[type="text"], input:not([type])`;

    function getPaperRoot(scope = document) {
        for (const selector of PAPER_SELECTORS) {
            if (scope.matches?.(selector)) return scope;
            const found = scope.querySelector(selector);
            if (found) return found;
        }
        return scope === document ? document.body : scope;
    }

    function selectIncludingRoot(root, selector) {
        return [...(root.matches?.(selector) ? [root] : []), ...root.querySelectorAll(selector)];
    }

    function getQuestions(root) {
        return selectIncludingRoot(root, QUESTION_SELECTOR).filter(el => !el.querySelector(QUESTION_SELECTOR));
    }

    function getStem(question) {
        for (const selector of STEM_SELECTOR.split(', ')) {
            const stem = question.matches(selector) ? question : question.querySelector(selector);
            if (stem && !stem.closest(ANSWER_SELECTOR)) return stem;
        }
        return question;
    }

    function hasPaperContent(el) {
        return el.matches(`h1, .report-title-a, ${QUESTION_SELECTOR}, ${STEM_SELECTOR}`) ||
            !!el.querySelector(`h1, .report-title-a, ${QUESTION_SELECTOR}, ${STEM_SELECTOR}`) ||
            PAPER_SELECTORS.some(selector => el.matches(selector) || el.querySelector(selector));
    }

    function escapeXML(text) {
        return String(text).replace(/[&<>"']/g, ch => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;'}[ch]));
    }

    function getPaperTitle() {
        const paper = getPaperRoot();
        for (const scope of [paper, document]) {
            for (const selector of ['h1', '.report-title-a', '.paper-title', '.report-title']) {
                const el = scope.querySelector(selector);
                if (el?.textContent.trim()) return el.textContent.trim();
            }
        }
        return (document.title || '菁优网试卷').replace(/[_-]\s*菁优网.*/, '').trim() || '菁优网试卷';
    }

    function removeFooters(scope = document) {
        const protectedArea = `${QUESTION_SELECTOR}, ${STEM_SELECTOR}, ${ANSWER_SELECTOR}, ${MATH_SELECTOR}, ${OPTION_SELECTOR}, #jyeoo-control-panel`;
        const canRemove = el => el !== scope && !hasPaperContent(el) &&
            !el.closest(protectedArea) && !el.matches('h1, h2, h3, h4, h5, h6, .report-title, .paper-title') &&
            !el.querySelector('h1, h2, h3, h4, h5, h6, .report-title, .paper-title');
        // 先识别整棵树，最后统一删除；防止先删版权段后丢失其余底栏的识别依据。
        const removals = new Set();
        scope.querySelectorAll('footer, .search-footer-2024, .footer, #Footer, #divFooter, [id$="Footer"], .report-info, .paper-info, .paper-sub-title, .paper-tool').forEach(el => {
            if (canRemove(el)) removals.add(el);
        });

        const compact = value => (value || '').replace(/\s+/g, '');
        const metaField = /(?:收录|总分|难度|浏览|下载)[：:]/g;
        const appField = /(?:APP开发者|应用名称|应用版本)[：:]/gi;
        const metaFragment = /(?:收录|总分|难度|浏览|下载)[：:]|显示答案|收藏试卷/;
        const footerFragment = /试卷征集|商务合作|服务条款|走进菁优|课件征集|帮助中心|兼职招聘|内容审核|菁优.*(?:教育|科技|公司)|ICP备|公网安备|jyeoo\.com|版权所有|市场监督|主体身份|APP开发者|应用名称|应用版本|隐私(?:政策|协议)|第三方SDK|广播电视|出版物经营|网站地图|本网站部分资源|版权归原作者|侵权|请(?:立即)?联系本站|菁优网在线客服/i;
        const childrenMatch = (el, predicate) => Array.from(el.childNodes).every(child => {
            if (child.nodeType === 8 || child.matches?.('script, style') || removals.has(child)) return true;
            const text = compact(child.textContent);
            return !text || /^[|｜·•—\-©\d:：/.()（）]+$/.test(text) || predicate(text);
        });
        const metadataChildren = value => metaFragment.test(value) || /^(?:易|中等|难|较易|较难)$/.test(value);
        const footerChildren = value => footerFragment.test(value) || /^(?:菁优网|本站|联系我们|举报)$/.test(value);
        scope.querySelectorAll('div, section, aside, footer, nav, ul, table, p, span').forEach(el => {
            const text = compact(el.textContent);
            if (!text || text.length > 6000) return;
            const fields = new Set((text.match(metaField) || []).map(field => field.slice(0, -1)));
            const appFields = new Set((text.match(appField) || []).map(field => field.slice(0, -1).toLowerCase()));
            const isMetadata = text.length < 600 && fields.size >= 3 &&
                ((text.includes('显示答案') && text.includes('收藏试卷')) || fields.size >= 4);
            // 下半部各段可独立出现，不再依赖已删掉的公司版权/备案段。
            const isAppInfo = text.length < 1400 && appFields.size === 3 && /菁优/.test(text) && /隐私(?:政策|协议)|第三方SDK|用户服务条款/.test(text);
            const isLicenses = /广播电视.*经营许可证/.test(text) && /出版物经营许可证/.test(text) && /网站地图/.test(text);
            const isResourceNotice = /本网站部分资源来源于会员上传/.test(text) && /版权归原作者所有/.test(text) && /侵权/.test(text);
            const isFooter = isAppInfo || isLicenses || isResourceNotice ||
                (text.includes('试卷征集') && text.includes('商务合作') && /菁优|jyeoo\.com/i.test(text)) ||
                (/jyeoo\.com/i.test(text) && /版权所有|copyright|©/i.test(text) && /ICP备|公网安备|隐私政策|服务条款|试卷征集|商务合作/.test(text));
            if ((!isMetadata && !isFooter) || !canRemove(el)) return;
            // 不删除混有独立备注/正文的外层，只收集纯信息栏/底栏容器。
            if ((isMetadata && childrenMatch(el, metadataChildren)) || (isFooter && childrenMatch(el, footerChildren))) removals.add(el);
        });
        removals.forEach(el => { if (scope.contains(el)) el.remove(); });
    }

    function removeWordTools(root) {
        const options = ['跨学科改编', '情景改编', '提升难度', '降低难度', '基础改编'];
        const fragments = /改编记录|【原题目】|帮我改编|改编目标|跨学科改编|情景改编|提升难度|降低难度|基础改编|关闭|取消|生成|确定/;
        const onlyToolContent = (node, depth = 0) => {
            if (depth > 40) return false;
            if (node.nodeType === 8) return true;
            if (node.nodeType === 3) return !node.textContent.trim() || fragments.test(node.textContent);
            if (node.matches?.('button, input, select, textarea')) return true;
            return Array.from(node.childNodes).every(child => onlyToolContent(child, depth + 1));
        };
        const removals = [];
        root.querySelectorAll('div, section, aside, dialog').forEach(el => {
            if (hasPaperContent(el) || el.closest(`${STEM_SELECTOR}, ${ANSWER_SELECTOR}, ${MATH_SELECTOR}, ${OPTION_SELECTOR}`)) return;
            const text = el.textContent.replace(/\s+/g, '');
            if (text.length > 4000 || !text.includes('改编记录') || !/帮我改编|改编目标/.test(text) || options.filter(value => text.includes(value)).length < 2) return;
            if (Array.from(el.childNodes).every(child => onlyToolContent(child))) removals.push(el);
        });
        removals.forEach(el => { if (root.contains(el)) el.remove(); });
    }

    function normalizeBlanks(root) {
        const questions = getQuestions(root);
        const stems = questions.length ? questions.map(getStem) : selectIncludingRoot(root, STEM_SELECTOR);
        const candidates = `${EXPLICIT_BLANK_SELECTOR}, .jyeoo-blank, .tunderline, .t-underline, .text-underline, u, span[style*="border-bottom"], font[style*="underline"]`;
        const seen = new Set();
        stems.forEach(stem => {
            selectIncludingRoot(stem, candidates).forEach(el => {
                if (seen.has(el) || !root.contains(el) || el.closest(`${ANSWER_SELECTOR}, ${OPTION_SELECTOR}`)) return;
                seen.add(el);
                if (el.classList.contains('jyeoo-blank')) return;
                const explicit = el.matches(EXPLICIT_BLANK_SELECTOR);
                // 站点真实填空可能包在 MathJye 中；明确占位优先，不误伤分式/根式线段。
                if (el.closest(MATH_SELECTOR) && (!explicit || el.closest('.mfrac, .msqrt, .munder, .mover, .msub, .msup, .msubsup'))) return;
                const empty = !el.textContent.replace(/[\s\u00a0_＿]/g, '') && !el.querySelector('img, table, input, svg');
                if (!explicit && !empty) return;
                const blank = document.createElement('u');
                blank.className = 'jyeoo-blank';
                if (el.matches(SITE_BLANK_SELECTOR)) blank.dataset.jyeooSiteBlank = 'true';
                let chars = el.textContent.replace(/[^\s\u00a0_＿]/g, '').length;
                const width = el.style.width || el.style.minWidth;
                if (/^[\d.]+em$/.test(width)) chars = Math.round(parseFloat(width) * 2);
                else if (/^[\d.]+px$/.test(width)) chars = Math.round(parseFloat(width) / 7);
                chars = Math.min(40, Math.max(4, chars || 10));
                blank.textContent = '\u00a0'.repeat(chars);
                blank.setAttribute('style', 'display:inline !important;text-decoration:underline !important;border:0 !important;visibility:visible !important;color:#000;white-space:pre;mso-spacerun:yes;');
                el.replaceWith(blank);
            });
        });
    }

    function cleanExtraSpaces(root) {
        if (!root) return;
        root.querySelectorAll(OPTION_SELECTOR).forEach(table => {
            table.style.setProperty('clear', 'none', 'important');
            table.style.setProperty('margin', '4px 0 2px', 'important');
        });
        root.querySelectorAll('.clear').forEach(el => {
            if (!el.textContent.trim() && !el.children.length) el.remove();
        });
        normalizeBlanks(root);
    }

    function processBlanksForWord(root) {
        if (!root) return;
        normalizeBlanks(root);
        root.querySelectorAll('li.QUES_LI, li.ques-item, li[id^="q_"]').forEach(li => {
            const div = document.createElement('div');
            Array.from(li.attributes).forEach(attr => div.setAttribute(attr.name, attr.value));
            div.append(...li.childNodes);
            li.replaceWith(div);
        });
        root.querySelectorAll(MATH_SELECTOR).forEach(el => {
            if (root.contains(el) && el.querySelector('.jyeoo-blank') && !el.querySelector('.mfrac, .msqrt, .msup, .msub, .msubsup, .munder, .mover, table, math, svg, img')) el.replaceWith(...el.childNodes);
        });
        const significantSibling = (node, direction) => {
            let sibling = node[direction];
            while (sibling && (sibling.nodeType === 8 || (sibling.nodeType === 3 && !sibling.textContent.trim()))) sibling = sibling[direction];
            return sibling;
        };
        const isInline = node => node && (node.nodeType === 3 || node.matches?.('span, u, b, strong, i, em, a, sup, sub, label'));
        root.querySelectorAll('.jyeoo-blank[data-jyeoo-site-blank]').forEach(blank => {
            let shell = blank.parentElement;
            while (shell && shell !== root && shell.matches('div, span, p') && !shell.matches(`${QUESTION_SELECTOR}, ${STEM_SELECTOR}`) &&
                   !shell.closest(`${ANSWER_SELECTOR}, table, ${MATH_SELECTOR}`) &&
                   Array.from(shell.childNodes).every(node => node === blank || node.nodeType === 8 || (node.nodeType === 3 && !node.textContent.trim()) || node.nodeName === 'BR')) {
                shell.replaceWith(blank); shell = blank.parentElement;
            }
            // p 内嵌填空 div 经 HTML 解析后常成为“句子 p + 填空”；仅合并这一类尾随 br 的断句。
            const previous = significantSibling(blank, 'previousSibling');
            if (previous?.matches?.('p') && !previous.querySelector('table, img') && significantSibling({ previousSibling: previous.lastChild }, 'previousSibling')?.nodeName === 'BR') {
                const inline = document.createElement('span');
                Array.from(previous.attributes).forEach(attr => inline.setAttribute(attr.name, attr.value));
                inline.append(...previous.childNodes);
                while (inline.lastChild && (inline.lastChild.nodeName === 'BR' || (inline.lastChild.nodeType === 3 && !inline.lastChild.textContent.trim()))) inline.lastChild.remove();
                inline.style.setProperty('display', 'inline', 'important'); previous.replaceWith(inline);
            }
            const breaks = [];
            let before = significantSibling(blank, 'previousSibling'), after = significantSibling(blank, 'nextSibling');
            while (before?.nodeName === 'BR') { breaks.push(before); before = significantSibling(before, 'previousSibling'); }
            while (after?.nodeName === 'BR') { breaks.push(after); after = significantSibling(after, 'nextSibling'); }
            if (isInline(before) && isInline(after)) breaks.forEach(br => br.remove());
        });
        root.querySelectorAll('div, p').forEach(block => {
            if (block.closest(`${ANSWER_SELECTOR}, table, ${MATH_SELECTOR}`) || block.attributes.length) return;
            if (block.childNodes.length === 1 && block.firstElementChild?.classList.contains('jyeoo-blank')) block.replaceWith(block.firstElementChild);
        });
    }

    // ================= 3. 有限 LaTeX/MathJye → OMML（失败时保留原式） =================
    const MATH_SYMBOLS = Object.freeze({
        times: '×', div: '÷', cdot: '·', leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈', pm: '±',
        rho: 'ρ', Omega: 'Ω', mu: 'μ', alpha: 'α', beta: 'β', lambda: 'λ', Delta: 'Δ', delta: 'δ', pi: 'π',
        theta: 'θ', omega: 'ω', gamma: 'γ', sigma: 'σ', eta: 'η', epsilon: 'ε', rightarrow: '→', circ: '°', infty: '∞',
        '%': '%', '{': '{', '}': '}', '_': '_', '&': '&', '#': '#', ' ': ' ', ',': ' ', ';': ' ', '!': ''
    });

    function mathText(text) {
        return `<m:r><m:t xml:space="preserve">${escapeXML(text)}</m:t></m:r>`;
    }

    function scriptOMML(base, sub, sup) {
        if (sub && sup) return `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
        if (sub) return `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
        return `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
    }

    function radicalOMML(value, degree = '') {
        return `<m:rad><m:radPr><m:degHide m:val="${degree ? 0 : 1}"></m:degHide></m:radPr><m:deg>${degree}</m:deg><m:e>${value}</m:e></m:rad>`;
    }

    function stripMathDelimiters(latex) {
        return latex.trim().replace(/^\$\$([\s\S]*)\$\$$/, '$1').replace(/^\$([\s\S]*)\$$/, '$1').replace(/^\\\(([\s\S]*)\\\)$/, '$1');
    }

    function latexToOMML(latex) {
        const input = stripMathDelimiters(latex);
        if (!input || input.length > 12000) throw new Error('公式为空或过长');
        let pos = 0, depth = 0;
        const skipSpaces = () => { while (/\s/.test(input[pos] || '') && pos < input.length) pos++; };
        function requiredGroup() {
            skipSpaces();
            if (input[pos] !== '{') throw new Error('公式参数缺少花括号');
            pos++;
            return sequence('}');
        }
        function atom(scriptArgument = false) {
            skipSpaces();
            const ch = input[pos++];
            if (!ch || /[\^_}\]]/.test(ch)) throw new Error('公式缺少基底或参数');
            if (ch === '{') return sequence('}');
            if (!scriptArgument && /\d/.test(ch)) {
                const number = input.slice(pos - 1).match(/^\d+(?:\.\d+)?/)[0];
                pos += number.length - 1;
                return mathText(number);
            }
            if (ch !== '\\') return mathText(ch);
            const match = input.slice(pos).match(/^[A-Za-z]+/);
            const command = match ? match[0] : input[pos];
            if (!command) throw new Error('公式命令不完整');
            pos += command.length;
            if (command === 'frac' || command === 'dfrac' || command === 'tfrac') {
                const num = requiredGroup(), den = requiredGroup();
                if (!num || !den) throw new Error('分式缺少分子或分母');
                return `<m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
            }
            if (command === 'sqrt') {
                skipSpaces();
                let degree = '';
                if (input[pos] === '[') { pos++; degree = sequence(']'); }
                const value = requiredGroup();
                if (!value) throw new Error('根式缺少被开方数');
                return radicalOMML(value, degree);
            }
            if (command === 'text' || command === 'mathrm') {
                skipSpaces();
                if (input[pos++] !== '{') throw new Error('文字命令缺少参数');
                const start = pos;
                while (pos < input.length && input[pos] !== '}') {
                    if (/[{\\]/.test(input[pos])) throw new Error('文字命令包含待支持结构');
                    pos++;
                }
                if (pos >= input.length) throw new Error('文字命令未闭合');
                return mathText(input.slice(start, pos++));
            }
            if (Object.prototype.hasOwnProperty.call(MATH_SYMBOLS, command)) return mathText(MATH_SYMBOLS[command]);
            throw new Error(`待支持命令 \\${command}`);
        }
        function sequence(end = '') {
            if (++depth > 64) throw new Error('公式嵌套过深');
            const nodes = [];
            while (pos < input.length) {
                skipSpaces();
                if (end && input[pos] === end) { pos++; depth--; return nodes.join(''); }
                if (pos >= input.length) break;
                let base = atom(), sub = '', sup = '';
                skipSpaces();
                while (input[pos] === '^' || input[pos] === '_') {
                    const op = input[pos++];
                    const value = atom(true);
                    if (!value || (op === '^' ? sup : sub)) throw new Error('重复或空上下标');
                    if (op === '^') sup = value; else sub = value;
                    skipSpaces();
                }
                if (sub || sup) base = scriptOMML(base, sub, sup);
                nodes.push(base);
            }
            depth--;
            if (end) throw new Error('公式分组未闭合');
            return nodes.join('');
        }
        return `<m:oMath>${sequence()}</m:oMath>`;
    }

    function domToOMML(node, depth = 0) {
        if (depth > 64) throw new Error('公式嵌套过深');
        if (node.nodeType === 3) return node.textContent.trim() ? mathText(node.textContent) : '';
        if (node.nodeType !== 1) return '';
        const children = Array.from(node.childNodes).filter(child => child.nodeType === 1 || child.textContent.trim());
        const convert = child => domToOMML(child, depth + 1);
        if (node.matches('.mfrac')) {
            const line = children.findIndex(child => child.nodeType === 1 && child.matches('.fracLine'));
            if (line < 1 || line === children.length - 1) throw new Error('分式 DOM 结构待支持');
            return `<m:f><m:num>${children.slice(0, line).map(convert).join('')}</m:num><m:den>${children.slice(line + 1).map(convert).join('')}</m:den></m:f>`;
        }
        if (node.matches('.msup, .msub, .msubsup')) {
            const expected = node.matches('.msubsup') ? 3 : 2;
            if (children.length !== expected) throw new Error('上下标 DOM 结构待支持');
            const parts = children.map(convert);
            return scriptOMML(parts[0], node.matches('.msub, .msubsup') ? parts[1] : '', node.matches('.msubsup') ? parts[2] : node.matches('.msup') ? parts[1] : '');
        }
        if (node.matches('.msqrt')) {
            const value = children.filter(child => !(child.nodeType === 1 && child.matches('.sqrt-symbol, .root-sign'))).map(convert).join('');
            if (!value) throw new Error('根式 DOM 内容为空');
            return radicalOMML(value);
        }
        if (node.matches('img, svg, math, .munder, .mover, .mtable, .mroot, .fracLine')) throw new Error('公式 DOM 结构待支持');
        // 未识别的结构类不扁平化为貌似正确的公式。
        if (Array.from(node.classList).some(name => /^m(?:frac|sqrt|sup|sub|under|over|table|root|atrix)/i.test(name))) throw new Error('公式 DOM 结构待支持');
        return children.map(convert).join('');
    }

    function simpleMathText(el, latex) {
        const fragment = document.createElement('span');
        const raw = el.textContent.trim();
        const unit = /^(?:\d+(?:\.\d+)?\s*)?(?:mm|cm|dm|m|km|mg|kg|g|ms|s|min|h|N|Pa|J|W|A|V|℃|m\/s|km\/h)$/;
        if ((!latex || !/[\\^_{}=<>]/.test(latex)) && !el.querySelector('.msup, .msub, .mfrac, .msqrt') && unit.test(raw)) {
            fragment.textContent = raw;
            return fragment;
        }
        const plain = stripMathDelimiters(latex || '').match(/^(10|mm|cm|dm|m|km)\^(?:\{(-?\d+)\}|(-?\d+))$/);
        if (plain && (plain[1] === '10' || /^[23]$/.test(plain[2] || plain[3]))) {
            fragment.textContent = plain[1];
            const sup = document.createElement('sup'); sup.textContent = plain[2] || plain[3]; fragment.append(sup);
            return fragment;
        }
        // 有明确 LaTeX 时不以可能已丢失结构的 rawText 推断角标。
        const simpleSub = stripMathDelimiters(latex || '').match(/^([A-Za-z])_(?:\{([A-Za-z0-9\u4e00-\u9fff]+)\}|([A-Za-z0-9\u4e00-\u9fff]))$/);
        if (simpleSub) {
            fragment.textContent = simpleSub[1];
            const sub = document.createElement('sub'); sub.textContent = simpleSub[2] || simpleSub[3]; fragment.append(sub);
            return fragment;
        }
        if (latex) return null;
        let script = el.matches('.msub, .msup') ? el : el.children.length === 1 && !Array.from(el.childNodes).some(n => n.nodeType === 3 && n.textContent.trim()) ? el.firstElementChild : null;
        if (!script?.matches('.msub, .msup') || script.children.length !== 2 || script.querySelector('.mfrac, .msqrt, .msub, .msup')) return null;
        const [base, value] = script.children;
        if (!base.textContent.trim() || !value.textContent.trim()) return null;
        fragment.textContent = base.textContent;
        const tag = document.createElement(script.matches('.msub') ? 'sub' : 'sup'); tag.textContent = value.textContent; fragment.append(tag);
        return fragment;
    }

    function processMathForWord(root, issues = []) {
        const entries = [];
        // 先快照顶层公式，避免父节点替换后又处理已脱离文档的子公式。
        const formulas = Array.from(root.querySelectorAll(MATH_SELECTOR)).filter(el => !el.parentElement?.closest(MATH_SELECTOR));
        formulas.forEach((el, index) => {
            const explicit = ['lazylatex', 'data-latex', 'mathtag'].map(name => el.getAttribute(name)).find(value => value?.trim());
            const title = el.getAttribute('title') || '';
            const latex = explicit || (/^(?:\$|\\\()/.test(title) ? title : '');
            if (el.querySelector('.jyeoo-blank')) { issues.push('含题空公式保留原布局，请核对显示效果'); return; }
            const simple = simpleMathText(el, latex);
            if (simple) { el.replaceWith(simple); return; }
            try {
                const xml = latex ? latexToOMML(latex) : `<m:oMath>${domToOMML(el)}</m:oMath>`;
                if (!xml.includes('<m:t')) throw new Error('公式内容为空');
                const parsed = new DOMParser().parseFromString(`<root xmlns:m="http://schemas.microsoft.com/office/2004/12/omml">${xml}</root>`, 'application/xml');
                if (parsed.querySelector('parsererror')) throw new Error('公式 XML 校验失败');
                const placeholder = document.createElement('span');
                placeholder.setAttribute('data-jyeoo-omml', String(index));
                placeholder.textContent = `JYEOO_OMML_${index}`;
                entries.push({ token: placeholder.outerHTML, xml });
                el.replaceWith(placeholder);
            } catch (error) {
                issues.push(`公式保留原式：${error.message}（${(latex || el.textContent).slice(0, 80)}）`);
                if (latex) {
                    const fallback = document.createElement('span'); fallback.textContent = latex; el.replaceWith(fallback);
                } // 无 LaTeX 时保留原有 DOM，避免扁平化复杂结构。
            }
        });
        return entries;
    }

    function serializeMath(root, entries) {
        let html = root.innerHTML;
        for (const entry of entries) {
            if (html.split(entry.token).length !== 2) throw new Error('公式占位符匹配异常，导出已停止');
            html = html.replace(entry.token, () => entry.xml);
        }
        return html;
    }

    // ================= 4. Word 图片与表格版面 =================
    function getImageSource(img) {
        const lazy = ['data-src', 'data-original', 'data-lazy-src'].map(key => img.getAttribute(key)).find(value => value?.trim());
        const raw = lazy || img.getAttribute('src') || '';
        if (!raw.trim()) return '';
        try {
            const url = new URL(raw, document.baseURI);
            return /^(https?:|data:|blob:)$/.test(url.protocol) ? url.href : '';
        } catch { return ''; }
    }

    function captureImages(original, clone) {
        const source = Array.from(original.querySelectorAll('img'));
        clone.querySelectorAll('img').forEach((img, index) => {
            const live = source[index], rect = live.getBoundingClientRect();
            const size = [rect.width, rect.height], natural = [live.naturalWidth, live.naturalHeight];
            const attrs = [Number(live.getAttribute('width')), Number(live.getAttribute('height'))];
            const dims = size.every(n => n > 0) ? size : attrs.every(n => n > 0) ? attrs : natural;
            if (dims.every(n => n > 0)) { img.dataset.jyeooWidth = String(dims[0]); img.dataset.jyeooHeight = String(dims[1]); }
            const boundary = live.closest(STEM_SELECTOR) || live.closest(QUESTION_SELECTOR) || original;
            let depth = 0;
            for (let el = live; el && el !== boundary && original.contains(el); el = el.parentElement, depth++) {
                const side = getComputedStyle(el).float || '';
                const legacy = el === live ? el.getAttribute('align') : '';
                if (['left', 'right'].includes(side) || ['left', 'right'].includes(legacy)) {
                    img.dataset.jyeooFloat = ['left', 'right'].includes(side) ? side : legacy;
                    img.dataset.jyeooFloatDepth = String(depth); break;
                }
            }
            const align = live.parentElement ? getComputedStyle(live.parentElement).textAlign : '';
            if (['center', 'right', 'left'].includes(align)) img.dataset.jyeooAlign = align;
            const src = getImageSource(live);
            if (src) img.setAttribute('src', src);
        });
    }

    function sizeWordImage(img, maxWidth, maxHeight = Infinity, issues = []) {
        const width = Number(img.dataset.jyeooWidth) || Number(img.getAttribute('width'));
        const height = Number(img.dataset.jyeooHeight) || Number(img.getAttribute('height'));
        if (!(width > 0 && height > 0)) { issues.push(`图片缺少宽高，保留原图：${(img.getAttribute('src') || '').slice(0, 100)}`); return; }
        const scale = Math.min(1, maxWidth / width, maxHeight / height);
        const w = Math.max(1, Math.round(width * scale)), h = Math.max(1, Math.round(height * scale));
        img.setAttribute('width', String(w)); img.setAttribute('height', String(h));
        img.style.setProperty('width', `${w}px`, 'important'); img.style.setProperty('height', `${h}px`, 'important');
        img.style.removeProperty('max-width'); img.style.removeProperty('max-height');
        // 浮动只在实际生成图文布局表格时清除；普通图片在此仅调整尺寸。
        img.dataset.jyeooSized = 'true';
    }

    function formatLayoutForWord(root, issues = []) {
        function floatCarrier(img, boundary) {
            const pack = (carrier, side) => {
                while (carrier.parentElement && carrier.parentElement !== boundary && carrier.parentElement.matches('a, span') && !carrier.parentElement.textContent.trim()) carrier = carrier.parentElement;
                return { carrier, side };
            };
            if (img.dataset.jyeooFloat) {
                let carrier = img;
                for (let depth = Number(img.dataset.jyeooFloatDepth) || 0; depth > 0 && carrier.parentElement !== boundary; depth--) carrier = carrier.parentElement;
                return pack(carrier, img.dataset.jyeooFloat);
            }
            for (let el = img; el && el !== boundary; el = el.parentElement) {
                const side = el.style.float || (el === img ? el.getAttribute('align') : '');
                if (side === 'left' || side === 'right') return pack(el, side);
            }
            return null;
        }
        getQuestions(root).forEach(question => {
            const stem = getStem(question);
            const images = Array.from(question.querySelectorAll('img')).filter(img => !img.closest(`${ANSWER_SELECTOR}, ${OPTION_SELECTOR}, ${MATH_SELECTOR}`));
            images.forEach(img => {
                if (img.closest('.jyeoo-stem-layout')) return;
                const boundary = stem.contains(img) ? stem : question;
                const floating = floatCarrier(img, boundary);
                if (!floating) {
                    if (['center', 'right'].includes(img.dataset.jyeooAlign)) {
                        let parent = img.parentElement;
                        while (parent && parent !== boundary && parent.matches('a, span')) parent = parent.parentElement;
                        if (parent) parent.style.textAlign = img.dataset.jyeooAlign;
                    }
                    return;
                }
                const { carrier } = floating;
                let host = carrier.parentElement;
                while (host && host !== boundary && !host.matches('div, p, section, article')) host = host.parentElement;
                if (!host || !boundary.contains(host) || carrier.querySelector(`${STEM_SELECTOR}, ${ANSWER_SELECTOR}, ${OPTION_SELECTOR}`)) return;
                let nodes = Array.from(host.childNodes).filter(node => !node.matches?.(`${ANSWER_SELECTOR}, ${OPTION_SELECTOR}`));
                if (host === question && stem !== question && !stem.contains(carrier)) {
                    nodes = Array.from(host.childNodes).filter(node => node === stem || node === carrier);
                } else if (host === boundary) {
                    const start = nodes.indexOf(carrier);
                    const priorParagraph = nodes.slice(0, start).some(node => node.matches?.('p, div, section, table'));
                    if (priorParagraph && start >= 0) {
                        let end = start + 1;
                        while (end < nodes.length) {
                            const node = nodes[end++];
                            if (node.matches?.('p, div, section, table')) break;
                        }
                        nodes = nodes.slice(start, end);
                    }
                }
                if (!nodes.length || !nodes.includes(carrier)) return;
                if (host.tagName === 'P') {
                    const div = document.createElement('div');
                    Array.from(host.attributes).forEach(attr => div.setAttribute(attr.name, attr.value));
                    div.append(...host.childNodes); host.replaceWith(div); host = div;
                }
                const groups = [floating];
                images.forEach(other => {
                    if (other === img || other.closest('.jyeoo-stem-layout')) return;
                    const candidate = floatCarrier(other, boundary);
                    if (candidate && nodes.includes(candidate.carrier) && !groups.some(group => group.carrier === candidate.carrier)) groups.push(candidate);
                });
                const optionTables = nodes.flatMap(node => node.querySelectorAll ? Array.from(node.querySelectorAll(OPTION_SELECTOR)) : [])
                    .filter(option => !option.parentElement.closest(OPTION_SELECTOR));
                optionTables.forEach(option => option.remove());
                const table = document.createElement('table'); table.className = 'jyeoo-stem-layout';
                table.setAttribute('border', '0'); table.setAttribute('cellspacing', '0'); table.setAttribute('cellpadding', '0');
                table.style.cssText = 'width:100%;border:none;border-collapse:collapse;margin:0 0 4pt;';
                const row = table.insertRow(), cells = {};
                if (groups.some(group => group.side === 'left')) cells.left = row.insertCell();
                const text = row.insertCell(); text.style.cssText = 'border:none;padding:0 8pt;vertical-align:top;text-align:left;';
                if (groups.some(group => group.side === 'right')) cells.right = row.insertCell();
                host.insertBefore(table, nodes[0]);
                groups.forEach(group => {
                    const cell = cells[group.side];
                    cell.style.cssText = `border:none;padding:0;vertical-align:top;text-align:${group.side};`;
                    const diagrams = group.carrier.matches('img') ? [group.carrier] : Array.from(group.carrier.querySelectorAll('img'));
                    diagrams.forEach(diagram => { sizeWordImage(diagram, 220, Infinity, issues); diagram.style.setProperty('float', 'none', 'important'); diagram.removeAttribute('align'); });
                    group.carrier.style.setProperty('float', 'none', 'important');
                    const width = Math.max(1, ...diagrams.map(diagram => Number(diagram.getAttribute('width')) || 220));
                    cell.setAttribute('width', String(Math.max(Number(cell.getAttribute('width')) || 0, width)));
                    cell.style.width = `${cell.getAttribute('width')}px`;
                    if (!group.carrier.matches('img')) group.carrier.style.width = `${width}px`;
                    cell.append(group.carrier);
                });
                text.append(...nodes.filter(node => node.parentNode === host));
                table.after(...optionTables);
            });
        });
        root.querySelectorAll(OPTION_SELECTOR).forEach(table => {
            const rows = Array.from(table.rows).filter(row => row.closest('table') === table);
            const complex = !!table.querySelector('table') || rows.some(row => Array.from(row.cells).some(cell => cell.colSpan !== 1 || cell.rowSpan !== 1));
            if (complex) return;
            table.setAttribute('border', '0'); table.style.cssText = 'width:100%;border:none;border-collapse:collapse;margin:4pt 0;';
            rows.forEach(row => Array.from(row.cells).forEach(cell => {
                cell.style.cssText = 'border:none;vertical-align:middle;text-align:left;padding:3pt 4pt;font-size:10pt;';
                if (row.cells.length === 4) { cell.setAttribute('width', '25%'); cell.style.width = '25%'; }
                cell.querySelectorAll('img').forEach(img => sizeWordImage(img, 100, 95, issues));
            }));
        });
    }


    // ================= 6. 注入终极强力净化样式表 =================
    function markSiteUI() {
        const selectors = '.top, #divNav, .report-info, .paper-info, .bshare-custom, .paper-tab-ctrl, .tab, .report-title-b, .paper-tool, #JYE-SEARCH-FORM, .search-hid, #pickerRegion, .sleft, .popover, .sfill, .sright, .qrcode-side, #divRight, .service, #divBread, .bread, .logo, .new-content-right, .paper-side-2024, #divRightBox, .pay-center, .box-overlay-pay, .down-pop, .fieldtip, #_message, .btm-nav, .new-fixed-bottom, .ques-opt, .sanweibar, .tools, .toolbar, .report-opt, .pt-opt, .jye-ad, .watermark, .fixed-box, .btn, .icon-opt';
        document.querySelectorAll(selectors).forEach(el => {
            if (!hasPaperContent(el) && !el.closest(`${STEM_SELECTOR}, ${ANSWER_SELECTOR}, ${MATH_SELECTOR}, #jyeoo-control-panel`)) el.dataset.jyeooUi = 'true';
        });
    }

    function injectPurifyCSS(includeAnswer) {
        let oldStyle = document.getElementById('jyeoo-purify-style');
        if (oldStyle) oldStyle.remove();
        markSiteUI();

        const printStyle = document.createElement('style');
        printStyle.id = 'jyeoo-purify-style';
        printStyle.innerHTML = `
            /* 仅隐藏已通过结构检查的站点 UI。 */
            [data-jyeoo-ui="true"] {
                display: none !important;
                height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
            }

            /* 保持控制面板在屏幕上始终强力置顶可见 */
            #jyeoo-control-panel {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                z-index: 2147483647 !important;
            }
            #jyeoo-control-panel * {
                visibility: visible !important;
                opacity: 1 !important;
            }

            /* 2.5 强力保护大标题居中显示 */
            h1, .report-title-a, .report-title h1 {
                display: block !important;
                visibility: visible !important;
                text-align: center !important;
                font-size: 24px !important;
                font-weight: bold !important;
                line-height: 1.4 !important;
                margin: 15px 0 10px 0 !important;
                padding: 0 !important;
                border: none !important;
                background: transparent !important;
            }

            /* 2. 页面主容器全宽与白底（覆盖 wrapper, content, new-content 等所有层级容器） */
            html, body, .wrapper, .content, .container, .main,
            .new-content, .report-main-box, .main-content,
            .new-content-left, .report-main, #divReportMain, .report-title,
            div[class*="content"], div[class*="wrapper"], div[class*="main"] {
                width: 100% !important;
                max-width: 100% !important;
                min-width: 0 !important;
                margin: 0 auto !important;
                padding: 0 !important;
                float: none !important;
                box-shadow: none !important;
                border: none !important;
                background: #fff !important;
                background-color: #fff !important;
                overflow: visible !important;
                box-sizing: border-box !important;
            }

            /* 3. 大题标题（如“一、选择题”）边距紧凑 */
            h2, h3, .h2-txt, .h2-txt-new, [class*="type-head"] {
                margin: 12px 0 4px 0 !important;
                padding: 0 !important;
                line-height: 1.5 !important;
            }

            /* 4. 彻底去框与紧凑题距：仅最外层题框保留一行空白（margin-bottom: 8px），子容器清零 */
            fieldset, .fieldset, .QUES_LI, .rpt_b, li.QUES_LI,
            .ques-detail, .quesborder, .ques-border {
                border: none !important;
                border-width: 0 !important;
                outline: none !important;
                box-shadow: none !important;
                border-radius: 0 !important;
                background: transparent !important;
                padding: 0 !important;
                margin: 0 0 8px 0 !important;
                cursor: default !important;
                display: block !important;
                visibility: visible !important;
            }

            /* 内部嵌套容器彻底清空边距，杜绝层层叠加 */
            .ques-item, .question-item, .paper-item, .report-item, .ques-con, .tb-paper,
            .pt0, .pt1, .pt2, .pt3, .pt4, .pt5, .pt6, .pt7, .pt8, .pt9,
            .q-item, .matblk, div[id^="q_"], div[id^="ques_"], li[id^="q_"], li[id^="ques_"] {
                border: none !important;
                border-width: 0 !important;
                outline: none !important;
                box-shadow: none !important;
                border-radius: 0 !important;
                background: transparent !important;
                padding: 0 !important;
                margin: 0 !important;
            }

            legend {
                display: none !important;
                border: none !important;
            }

            /* 5. 填空题下划线强力保护与显示 */
            .jyeoo-blank {
                display: inline !important;
                text-decoration: underline !important;
                border: none !important;
                vertical-align: baseline !important;
                visibility: visible !important;
                padding: 0 !important;
                margin: 0 !important;
            }

            /* MathJye 公式系统内的下划线线段与 munder 结构保护 */
            .MathJye .tline, .MathJye .tlineSingle, .MathJye .munder, .MathJye .fracLine {
                border-top-color: #000 !important;
                border-bottom-color: #000 !important;
                visibility: visible !important;
                opacity: 1 !important;
            }

            /* 6. 选项表格取消 clear:both，紧跟题干右侧插图 */
            table.composition, table.composition2, .composition {
                clear: none !important;
                margin: 4px 0 2px 0 !important;
            }

            /* 7. 解析显隐控制 */
            ${ANSWER_SELECTOR} {
                display: ${includeAnswer ? 'block' : 'none'} !important;
                ${includeAnswer ? 'height:auto !important; visibility:visible !important;' : 'height:0 !important; margin:0 !important; padding:0 !important;'}
            }
            /* 8. 打印分页与自然接排控制，统一精确设定字号与版面全宽 */
            @media print {
                @page {
                    size: A4 portrait;
                    margin: 12mm 15mm 12mm 15mm;
                }
                html, body {
                    font-size: ${(PDF_TYPOGRAPHY.bodyPt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                }

                #jyeoo-control-panel {
                    display: none !important;
                }

                /* 彻底抹平 1200px 等固定宽，撑满 A4 页面版心，左右完全平衡 */
                html, body, .wrapper, .content, .container, .main,
                .new-content, .report-main-box, .main-content,
                .new-content-left, .report-main, #divReportMain,
                .ques-list, ul, ol, [class*="list" i], [class*="main" i],
                div[class*="content"], div[class*="wrapper"] {
                    width: 100% !important;
                    max-width: 100% !important;
                    min-width: 0 !important;
                    margin: 0 auto !important;
                    padding: 0 !important;
                    float: none !important;
                    box-sizing: border-box !important;
                    page-break-inside: auto !important;
                    break-inside: auto !important;
                    page-break-before: avoid !important;
                    break-before: avoid !important;
                }

                fieldset, .fieldset, .QUES_LI, .rpt_b, li.QUES_LI,
                .ques-detail, .quesborder, .ques-item, .question-item, .ques-con,
                table.composition, table.composition2, .composition {
                    width: 100% !important;
                    max-width: 100% !important;
                    box-sizing: border-box !important;
                }

                /* 试卷大标题：最终 PDF 目标 10pt（含实测缩放补偿） */
                h1, .report-title-a, .report-title h1 {
                    font-size: ${(PDF_TYPOGRAPHY.titlePt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                    font-weight: bold !important;
                    text-align: center !important;
                    margin: 8pt 0 6pt 0 !important;
                    page-break-after: avoid !important;
                    break-after: avoid !important;
                }

                /* 大题标题：最终 PDF 目标 10pt，加粗 */
                h2, h3, .h2-txt, .h2-txt-new, [class*="type-head"] {
                    font-size: ${(PDF_TYPOGRAPHY.titlePt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                    font-weight: bold !important;
                    margin: 10pt 0 4pt 0 !important;
                    padding: 0 !important;
                    line-height: 1.5 !important;
                    page-break-after: avoid !important;
                    break-after: avoid !important;
                    page-break-before: auto !important;
                    break-before: auto !important;
                }

                /* 题干、选项与正文：最终 PDF 目标 9pt（含实测缩放补偿） */
                :is(fieldset, .fieldset, .QUES_LI, .rpt_b, li.QUES_LI,
                .ques-detail, .quesborder, .ques-item, .question-item, .ques-con,
                .pt0, .pt1, .pt2, .pt3, .pt4, .pt5, .pt6, .pt7, .pt8, .pt9,
                .q-item, .matblk, div[id^="q_"], div[id^="ques_"], li[id^="q_"], li[id^="ques_"],
                p, span:not(sup *):not(sub *), table, td, th, label):not(:is(${MATH_SELECTOR})):not(:is(${MATH_SELECTOR}) *) {
                    font-size: ${(PDF_TYPOGRAPHY.bodyPt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                    line-height: 1.5 !important;
                }
                /* 公式只设置顶层字号；保留 MathJye 的零行高及内部节点各自的专用行高。
                   不能对 .MathJye * 统一设置 normal，否则仍会破坏根号伸缩与分式定位。 */
                :is(${MATH_SELECTOR}):not(:is(${MATH_SELECTOR}) *) {
                    font-size: ${(PDF_TYPOGRAPHY.bodyPt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                }

                /* 仅对单道具体试题做防截断保护 */
                fieldset, .QUES_LI, .rpt_b, .ques-detail, .quesborder, li.QUES_LI {
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                    page-break-before: auto !important;
                    break-before: auto !important;
                    border: none !important;
                    margin-bottom: 8px !important;
                    padding: 0 !important;
                }

                .jyeoo-blank {
                    display: inline !important;
                    text-decoration: underline !important;
                    border: none !important;
                }
                img {
                    max-width: 100% !important;
                }
                /* 标题及内部文字显式覆盖通用 span/p 字号；不覆盖公式及其内部相对尺寸。 */
                :is(h1, h2, h3, h4, h5, h6, .report-title-a, .h2-txt, .h2-txt-new, [class*="type-head"]):not(:is(${MATH_SELECTOR})):not(:is(${MATH_SELECTOR}) *),
                :is(h1, h2, h3, h4, h5, h6, .report-title-a, .h2-txt, .h2-txt-new, [class*="type-head"]) :is(p, span, label):not(:is(${MATH_SELECTOR})):not(:is(${MATH_SELECTOR}) *):not(sup *):not(sub *) {
                    font-size: ${(PDF_TYPOGRAPHY.titlePt / PDF_TYPOGRAPHY.measuredScale).toFixed(4)}pt !important;
                }
                [data-jyeoo-long="true"] { break-inside: auto !important; page-break-inside: auto !important; }
            }
        `;
        document.head.appendChild(printStyle);

        // JS 层深度清空题框外边框
        document.querySelectorAll(`
            fieldset, legend, .QUES_LI, .rpt_b, li.QUES_LI,
            .ques-detail, .quesborder, .ques-border, .ques-item,
            .pt0, .pt1, .pt2, .pt3, .pt4, .pt5, .pt6, .pt7, .pt8, .pt9,
            .matblk, div[id^="q_"], div[id^="ques_"], li[id^="q_"], li[id^="ques_"]
        `).forEach(el => {
            el.style.setProperty('border', 'none', 'important');
            el.style.setProperty('border-width', '0', 'important');
            el.style.setProperty('outline', 'none', 'important');
            el.style.setProperty('box-shadow', 'none', 'important');
            el.style.setProperty('border-radius', '0', 'important');
            el.style.setProperty('background', 'transparent', 'important');
        });

        // 执行空行与空白清理，全面保护填空题下划线
        cleanExtraSpaces(getPaperRoot());
        selectIncludingRoot(getPaperRoot(), QUESTION_SELECTOR).forEach(el => {
            el.toggleAttribute('data-jyeoo-long', el.getBoundingClientRect().height > 980);
            if (el.hasAttribute('data-jyeoo-long')) el.dataset.jyeooLong = 'true';
        });
    }

    // ================= 7. 导出 Word (.doc) 核心函数 =================
    function buildWordDocument(mainEl, includeAnswer, issues = []) {
        const title = getPaperTitle();
        const clone = mainEl.cloneNode(true);
        captureImages(mainEl, clone);
        removeFooters(clone);
        removeWordTools(clone);

        // 剔除无用节点
        const unwantedSelectors = [
            'header', 'footer', 'nav', 'script', 'style', 'iframe', '#jyeoo-control-panel', '[data-jyeoo-ui="true"]',
            '.top', '#divNav', '.report-info', '.paper-info', '.bshare-custom',
            '.paper-tab-ctrl', '.tab', '.report-title-b', '.paper-title-b', '.paper-tool', '.paper-sub-title',
            '#JYE-SEARCH-FORM', '.search-hid', '#pickerRegion', '.sleft', '.popover',
            '.qrcode-side', '#divRight', '.service', '#divBread', '.bread', '.logo',
            '.new-content-right', '.paper-side-2024', '#divRightBox', '.pay-center',
            '.down-pop', '.fieldtip', '#_message', '.btm-nav', '.new-fixed-bottom',
            '.jye-ad', '.ques-opt', '.sanweibar', '.tools', '.toolbar',
            '.btn', 'button', '.icon-opt', '.report-opt', '.pt-opt', '.share-box', '.watermark'
        ];
        unwantedSelectors.forEach(sel => {
            clone.querySelectorAll(sel).forEach(el => {
                if (el.matches('script, style, iframe, #jyeoo-control-panel') || (!hasPaperContent(el) && !el.closest(`${STEM_SELECTOR}, ${ANSWER_SELECTOR}, ${MATH_SELECTOR}`))) el.remove();
            });
        });
        // 移除克隆体内部可能存在的原标题，统一在最顶层由 <h1>${title}</h1> 渲染，避免标题重复
        clone.querySelectorAll('h1, .report-title-a').forEach(h1 => h1.remove());


        // 根据开关决定是否删除答案解析
        if (!includeAnswer) clone.querySelectorAll(ANSWER_SELECTOR).forEach(el => el.remove());
        else clone.querySelectorAll(ANSWER_SELECTOR).forEach(el => {
            el.removeAttribute('hidden');
            el.style.setProperty('display', 'block', 'important');
            el.style.setProperty('visibility', 'visible', 'important');
            el.style.removeProperty('height');
        });
        cleanExtraSpaces(clone);
        processBlanksForWord(clone);
        formatLayoutForWord(clone, issues);

        // 彻底清空克隆体内的题框边框，且每题之间只保留 8pt 紧凑间距
        clone.querySelectorAll(`
            fieldset, .QUES_LI, .rpt_b, .ques-detail, .quesborder, .ques-border
        `).forEach(el => {
            el.style.border = 'none';
            el.style.outline = 'none';
            el.style.boxShadow = 'none';
            el.style.borderRadius = '0';
            el.style.background = 'transparent';
            el.style.padding = '0';
            el.style.margin = '0 0 8pt 0';
        });

        clone.querySelectorAll(`
            .ques-item, .pt0, .pt1, .pt2, .pt3, .pt4, .pt5, .pt6, .pt7, .pt8, .pt9,
            .matblk, div[id^="q_"], div[id^="ques_"], li[id^="q_"], li[id^="ques_"]
        `).forEach(el => {
            el.style.border = 'none';
            el.style.outline = 'none';
            el.style.boxShadow = 'none';
            el.style.borderRadius = '0';
            el.style.background = 'transparent';
            el.style.padding = '0';
            el.style.margin = '0';
        });
        clone.querySelectorAll('legend').forEach(el => el.remove());

        clone.querySelectorAll('img').forEach(img => {
            const src = getImageSource(img);
            if (src) img.setAttribute('src', src);
            else { img.removeAttribute('src'); issues.push('图片地址缺失或格式待支持'); }
            if (!img.dataset.jyeooSized) sizeWordImage(img, 480, Infinity, issues);
            img.removeAttribute('srcset'); img.removeAttribute('loading');
            img.style.verticalAlign = 'middle';
            for (const attr of Array.from(img.attributes)) {
                if (attr.name.startsWith('data-') || attr.name.startsWith('on')) img.removeAttribute(attr.name);
            }
        });
        const formulas = processMathForWord(clone, issues);
        const bodyHtml = serializeMath(clone, formulas);
        if (!clone.textContent.trim() && !clone.querySelector('img')) throw new Error('未找到试卷正文，请等待页面加载完成');

        // 组装标准 Microsoft Word 结构（含 Office OMML 命名空间与 1.8cm 页面设置）
        const wordHtml = `
        <html xmlns:v="urn:schemas-microsoft-com:vml"
              xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <!--[if gte mso 9]>
        <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
          <w:Compatibility>
            <w:UnderlineTrailingSpaces/>
          </w:Compatibility>
        </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page Section1 {
            size: 595.3pt 841.9pt; /* A4 标准尺寸 */
            margin: 51.0pt 51.0pt 51.0pt 51.0pt; /* 上下左右各 1.8cm 标准紧凑对称页边距 */
            mso-header-margin: 36.0pt;
            mso-footer-margin: 36.0pt;
            mso-paper-source: 0;
          }
          div.Section1 {
            page: Section1;
            width: 100%;
          }
          body, fieldset, .quesborder, .QUES_LI, p, span, td, table {
            font-family: "SimSun", "宋体", "Times New Roman", serif;
            font-size: 10pt !important;
            line-height: 1.4 !important;
            color: #000000;
            background: #ffffff;
          }
          h1, .paper-title, .report-title {
            text-align: center;
            font-family: "SimHei", "黑体", sans-serif;
            font-size: 16pt !important;
            font-weight: bold;
            margin: 6pt 0 14pt 0 !important;
          }
          h2, h3, .h2-txt, .h2-txt-new {
            margin: 10pt 0 4pt 0 !important;
            font-family: "SimHei", "黑体", sans-serif;
            font-size: 11pt !important;
            font-weight: bold !important;
          }
          fieldset, .quesborder, .QUES_LI {
            border: none !important;
            margin-bottom: 8pt !important;
            padding: 0 !important;
          }
          legend {
            display: none !important;
          }
          .jyeoo-blank {
            display: inline !important;
            text-decoration: underline !important;
            border: none !important;
            vertical-align: baseline !important;
          }
          img {
            max-width: 100%;
            vertical-align: middle;
          }
          table { border-collapse: collapse; }
          .jyeoo-stem-layout, table.composition, table.composition2 { width: 100%; }
          .jyeoo-stem-layout td { border: none; }
          h1 span, h2 span, h3 span, .h2-txt span, .h2-txt-new span { font-size: inherit !important; }
          table.composition {
            margin: 4pt 0 !important;
          }
          table.composition td {
            font-size: 10pt !important;
            padding: 2pt 4pt !important;
          }
        </style>
        </head>
        <body>
          <div class="Section1">
            <h1>${escapeXML(title)}</h1>
            <div class="word-paper-body">
              ${bodyHtml}
            </div>
          </div>
        </body>
        </html>`;

        const safeTitle = title.replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/[. ]+$/, '').slice(0, 120) || '菁优网试卷';
        return { html: wordHtml, filename: `${safeTitle}${includeAnswer ? '（含解析）' : '（学生版）'}.doc` };
    }

    async function exportToWord(includeAnswer) {
        const mainEl = getPaperRoot(), issues = [];
        await waitForResources(mainEl, issues, 8000, includeAnswer);
        const result = buildWordDocument(mainEl, includeAnswer, issues);
        if (!confirmIssues(issues, 'Word 导出')) return;
        const blob = new Blob(['\ufeff' + result.html], { type: 'application/msword;charset=utf-8' });
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = result.filename;
        document.body.appendChild(link);
        try { link.click(); }
        finally { link.remove(); setTimeout(() => URL.revokeObjectURL(downloadUrl), 5000); }
    }

    // ================= 8. 资源就绪、导出状态与打印 =================
    async function waitForResources(root, issues = [], timeout = 8000, includeAnswer = document.getElementById('jyeoo-chk-answer')?.checked ?? true) {
        const controller = new AbortController(), signal = controller.signal;
        const pending = new Set(), jobs = [];
        root.querySelectorAll('img').forEach((img, index) => {
            if (img.closest(ANSWER_SELECTOR) && !includeAnswer) return;
            const label = `图片 ${index + 1}`, src = getImageSource(img);
            if (!src) { issues.push(`${label}地址缺失`); return; }
            if (img.src !== src) img.src = src;
            img.removeAttribute('loading');
            pending.add(label);
            jobs.push(new Promise(resolve => {
                let settled = false;
                const finish = ok => {
                    if (settled) return;
                    settled = true; pending.delete(label);
                    img.removeEventListener('load', loaded); img.removeEventListener('error', failed); signal.removeEventListener('abort', aborted);
                    if (!ok) issues.push(`${label}加载或解码未完成`);
                    resolve();
                };
                const loaded = () => {
                    if (typeof img.decode === 'function') img.decode().then(() => finish(true), () => finish(false));
                    else finish(img.naturalWidth > 0);
                };
                const failed = () => finish(false), aborted = () => finish(false);
                signal.addEventListener('abort', aborted, { once: true });
                img.addEventListener('load', loaded, { once: true }); img.addEventListener('error', failed, { once: true });
                if (img.complete) { if (img.naturalWidth > 0) loaded(); else failed(); }
            }));
        });
        if (document.fonts?.ready) {
            pending.add('字体');
            jobs.push(document.fonts.ready.then(() => pending.delete('字体'), () => { pending.delete('字体'); issues.push('字体加载未完成'); }));
        }
        let timer;
        try {
            await Promise.race([Promise.all(jobs), new Promise(resolve => { timer = setTimeout(resolve, timeout); })]);
            if (pending.has('字体')) issues.push('字体等待超时');
        } finally { clearTimeout(timer); controller.abort(); }
    }

    function confirmIssues(issues, action) {
        const unique = [...new Set(issues)];
        if (!unique.length) return true;
        console.warn(`[菁优网] ${action}`, unique);
        return window.confirm(`${action}存在 ${unique.length} 项待核对内容：\n${unique.slice(0, 8).join('\n')}\n\n继续导出保留现有内容；取消后可等待资源或检查原题。`);
    }

    let exportBusy = false;
    async function runExport(action) {
        if (exportBusy) return;
        exportBusy = true;
        const controls = Array.from(document.querySelectorAll('#jyeoo-control-panel button, #jyeoo-control-panel input'));
        const disabled = controls.map(el => el.disabled);
        controls.forEach(el => { el.disabled = true; });
        try { await action(); }
        catch (error) { console.error('[菁优网] 导出中止', error); window.alert(`导出中止：${error.message}`); }
        finally { controls.forEach((el, index) => { el.disabled = disabled[index]; }); exportBusy = false; }
    }

    async function exportToPDF() {
        const issues = [];
        await waitForResources(getPaperRoot(), issues);
        if (!confirmIssues(issues, 'PDF 打印')) return;
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        document.getElementById('jyeoo-control-panel')?.style.removeProperty('display');
        await new Promise((resolve, reject) => {
            let timer, settled = false;
            const done = () => {
                if (settled) return;
                settled = true; clearTimeout(timer);
                window.removeEventListener('afterprint', done); window.removeEventListener('focus', focused);
                resolve();
            };
            const focused = () => { if (!window.matchMedia('print').matches) done(); };
            window.addEventListener('afterprint', done, { once: true }); window.addEventListener('focus', focused);
            try { window.print(); if (!settled) timer = setTimeout(done, 60000); }
            catch (error) { window.removeEventListener('afterprint', done); window.removeEventListener('focus', focused); reject(error); }
        });
    }

    // ================= 9. 构建现代交互悬浮控制面板 =================
    const panel = document.createElement('div');
    panel.id = 'jyeoo-control-panel';
    panel.style.cssText = `
        position: fixed !important;
        bottom: 30px !important;
        right: 30px !important;
        z-index: 2147483647 !important;
        background: #ffffff !important;
        border: 2px solid #00b4d8 !important;
        border-radius: 14px !important;
        padding: 16px 18px !important;
        box-shadow: 0 10px 30px rgba(0,0,0,0.25) !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif !important;
        display: flex;
        flex-direction: column !important;
        gap: 12px !important;
        width: 250px !important;
        box-sizing: border-box !important;
    `;
    document.body.appendChild(panel);

    let isPurified = false;

    // 渲染面板UI（支持“未净化”与“净化后”两阶段无缝切换）
    function renderPanel() {
        const oldChk = document.getElementById('jyeoo-chk-answer');
        const includeAnswer = oldChk ? oldChk.checked : true;

        if (!isPurified) {
            // 阶段一：未净化前，提供大按钮【🚀 净化页面】
            panel.innerHTML = `
                <div style="font-size: 15px; font-weight: bold; color: #0077b6; display: flex; align-items: center; justify-content: center; gap: 6px; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">
                    <span style="font-size: 17px; line-height: 1;">📑</span>
                    <span>菁优网试卷提取器</span>
                </div>
                <label style="font-size: 14px; color: #333; display: flex; align-items: center; justify-content: center; cursor: pointer; user-select: none;">
                    <input type="checkbox" id="jyeoo-chk-answer" ${includeAnswer ? 'checked' : ''} style="margin-right: 8px; width: 16px; height: 16px; cursor: pointer;">
                    <span>包含答案与解析</span>
                </label>
                <button id="jyeoo-btn-purify" style="
                    width: 100% !important;
                    height: 42px !important;
                    box-sizing: border-box !important;
                    display: flex !important;
                    align-items: center !important;
                    justify-content: center !important;
                    gap: 8px !important;
                    font-size: 15px !important;
                    font-weight: bold !important;
                    white-space: nowrap !important;
                    color: #ffffff !important;
                    background: linear-gradient(135deg, #0077b6, #00b4d8) !important;
                    border: none !important;
                    border-radius: 8px !important;
                    cursor: pointer !important;
                    box-shadow: 0 4px 12px rgba(0,180,216,0.35) !important;
                    transition: all 0.2s !important;
                ">
                    <span style="font-size: 17px; line-height: 1;">🚀</span>
                    <span style="line-height: 1;">净化页面</span>
                </button>
            `;

            document.getElementById('jyeoo-btn-purify').addEventListener('click', () => {
                isPurified = true;
                const withAnswer = document.getElementById('jyeoo-chk-answer').checked;
                removeFooters();
                injectPurifyCSS(withAnswer);
                renderPanel();
            });

        } else {
            // 阶段二：净化后，停留在纯净试卷页面，面板提供【导出PDF】与【导出Word】
            panel.innerHTML = `
                <div style="font-size: 15px; font-weight: bold; color: #2b9348; display: flex; align-items: center; justify-content: center; gap: 6px; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">
                    <span style="font-size: 17px; line-height: 1;">✨</span>
                    <span>试卷已净化</span>
                </div>
                <label style="font-size: 14px; color: #333; display: flex; align-items: center; justify-content: center; cursor: pointer; user-select: none;">
                    <input type="checkbox" id="jyeoo-chk-answer" ${includeAnswer ? 'checked' : ''} style="margin-right: 8px; width: 16px; height: 16px; cursor: pointer;">
                    <span>包含答案与解析</span>
                </label>
                <button id="jyeoo-btn-pdf" style="
                    width: 100% !important;
                    height: 42px !important;
                    box-sizing: border-box !important;
                    display: flex !important;
                    align-items: center !important;
                    justify-content: center !important;
                    gap: 8px !important;
                    font-size: 14.5px !important;
                    font-weight: bold !important;
                    white-space: nowrap !important;
                    color: #ffffff !important;
                    background: linear-gradient(135deg, #0077b6, #00b4d8) !important;
                    border: none !important;
                    border-radius: 8px !important;
                    cursor: pointer !important;
                    box-shadow: 0 4px 10px rgba(0,180,216,0.3) !important;
                    transition: all 0.2s !important;
                ">
                    <span style="font-size: 17px; line-height: 1;">🖨️</span>
                    <span style="line-height: 1;">导出为 PDF</span>
                </button>
                <button id="jyeoo-btn-word" style="
                    width: 100% !important;
                    height: 42px !important;
                    box-sizing: border-box !important;
                    display: flex !important;
                    align-items: center !important;
                    justify-content: center !important;
                    gap: 8px !important;
                    font-size: 14.5px !important;
                    font-weight: bold !important;
                    white-space: nowrap !important;
                    color: #ffffff !important;
                    background: linear-gradient(135deg, #2b9348, #55a630) !important;
                    border: none !important;
                    border-radius: 8px !important;
                    cursor: pointer !important;
                    box-shadow: 0 4px 10px rgba(43,147,72,0.3) !important;
                    transition: all 0.2s !important;
                ">
                    <span style="font-size: 17px; line-height: 1;">📄</span>
                    <span style="line-height: 1;">导出为 Word (.doc)</span>
                </button>
            `;

            // 净化状态下切换勾选框，实时在当前页面展示或隐藏答案
            document.getElementById('jyeoo-chk-answer').addEventListener('change', (e) => {
                injectPurifyCSS(e.target.checked);
            });

            // 导出 PDF
            document.getElementById('jyeoo-btn-pdf').addEventListener('click', () => {
                runExport(() => exportToPDF());
            });

            // 导出 Word
            document.getElementById('jyeoo-btn-word').addEventListener('click', () => {
                const withAnswer = document.getElementById('jyeoo-chk-answer').checked;
                runExport(() => exportToWord(withAnswer));
            });
        }
    }

    renderPanel();

})();

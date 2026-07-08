(function () {
    const scriptUrl = document.currentScript && document.currentScript.src ? document.currentScript.src : "";
    const PDF_BASE = scriptUrl ? new URL("output/pdf/", scriptUrl).href : "output/pdf/";
    const PDF_MAPPING_URL = scriptUrl ? new URL("pdf-url-mapping.csv", scriptUrl).href : "pdf-url-mapping.csv";
    const PDF_URL_MAP = new Map();
    let pdfUrlMappingPromise = null;
    const PDF_FILES = [
        "MSF11-13 核查记录表（2015版）最新勘误.pdf",
        "MSF11-04 管理体系过程清单（2019.9.5修订）.pdf",
        "MSF11-26 审核人员商调函（2015版）.pdf",
        "MSF11-03 审核项目推进表（2018.12.7修订）.pdf",
        "MSF11-27+食品安全认证申请组织自我声明（2022.1.13修订）.pdf",
        "MSF11-08+第一阶段不实施现场审核的申请（2022.1.13修订）.pdf",
        "MSF11-02+认证申请评审表（2025.12.31修订）.pdf",
        "MSWM11-14-01+认证转换前评审记录（2018.6.13修订）.pdf",
        "MSF11-17+审核报告（2022.1.13修订）.pdf",
        "MSF11-18 证书内容确认件（2024.12.6修订）.pdf",
        "MSF11-19+认证发证推荐表（2022.1.13修订）.pdf",
        "MSF11-21+文件评审报告（2022.1.13修订）.pdf",
        "MSF11-22 监督审核通知（2015版）.pdf",
        "MSF11-14+首末次会议签到表（2025.12.31修订）.pdf",
        "MSF11-01-认证申请表-演示预览版.pdf",
        "MSF11-07+审核任务书（2025.12.31修订）.pdf",
        "MSF13-04+认证评定表（2024.3.28修订）.pdf",
        "MSF11-25+证书变更申请表（2025.12.31修订）.pdf",
        "MSF11-24 再认证通知书（2015版）.pdf",
        "MSF13-02+监督审核认证评定表（2015版）.pdf",
        "MSF11-15+不符合项报告（2025.12.31修订）.pdf",
        "MSF11-09+第一阶段审核报告（2023.3.30修订）.pdf",
        "MSF11-12 审核准备会记录（2015版）.pdf",
        "MSF108-10+知识产权管理体系审核材料归档上报清单（2024.3.15修订）.pdf",
        "MSF13-01+认证决定通知书（2015版）.pdf",
        "MSF11-28 出口食品生产企业自我评估表（2016.9.30修订）.pdf",
        "MSF11-05 认证场所清单（2024.7.12修订）.pdf",
        "MSWM11-21-01 特殊领域认证项目审批表（2025.3.14修订）.pdf",
        "MSF11-11+审核计划（2025.12.31修订）.pdf",
        "MSF13-03+监督审核合格通知书（2017.6.13修订）.pdf",
        "MSF11-23+获证方情况调查表（2025.12.31修订）.pdf",
        "MSF11-20+审核材料归档上报清单（2025.12.31修订）.pdf",
        "MSF11-16+认证周期绩效评价表（2025.12.31修订）.pdf",
        "MSF11-10 第一阶段审核问题清单（2015版）.pdf"
    ];

    const PDF_MATCHERS = [
        [/认证申请评审|申请评审/, "MSF11-02+认证申请评审表（2025.12.31修订）.pdf"],
        [/认证申请表|申请表/, "MSF11-01-认证申请表-演示预览版.pdf"],
        [/证书内容|证书确认|确认件/, "MSF11-18 证书内容确认件（2024.12.6修订）.pdf"],
        [/认证发证推荐|发证推荐/, "MSF11-19+认证发证推荐表（2022.1.13修订）.pdf"],
        [/获证方情况|情况调查/, "MSF11-23+获证方情况调查表（2025.12.31修订）.pdf"],
        [/证书变更/, "MSF11-25+证书变更申请表（2025.12.31修订）.pdf"],
        [/认证周期绩效|绩效评价/, "MSF11-16+认证周期绩效评价表（2025.12.31修订）.pdf"],
        [/管理体系过程|过程清单/, "MSF11-04 管理体系过程清单（2019.9.5修订）.pdf"],
        [/认证场所/, "MSF11-05 认证场所清单（2024.7.12修订）.pdf"],
        [/审核项目推进|推进表/, "MSF11-03 审核项目推进表（2018.12.7修订）.pdf"],
        [/审核任务书|任务书/, "MSF11-07+审核任务书（2025.12.31修订）.pdf"],
        [/一阶段不实施|第一阶段不实施/, "MSF11-08+第一阶段不实施现场审核的申请（2022.1.13修订）.pdf"],
        [/第一阶段审核问题|一阶段审核问题|问题清单/, "MSF11-10 第一阶段审核问题清单（2015版）.pdf"],
        [/第一阶段审核报告|一阶段审核报告/, "MSF11-09+第一阶段审核报告（2023.3.30修订）.pdf"],
        [/文件评审/, "MSF11-21+文件评审报告（2022.1.13修订）.pdf"],
        [/审核准备会/, "MSF11-12 审核准备会记录（2015版）.pdf"],
        [/首末次会议|签到表/, "MSF11-14+首末次会议签到表（2025.12.31修订）.pdf"],
        [/审核计划|计划书/, "MSF11-11+审核计划（2025.12.31修订）.pdf"],
        [/不符合项|纠正措施/, "MSF11-15+不符合项报告（2025.12.31修订）.pdf"],
        [/核查记录/, "MSF11-13 核查记录表（2015版）最新勘误.pdf"],
        [/审核报告|二阶段审核报告|第二阶段审核报告/, "MSF11-17+审核报告（2022.1.13修订）.pdf"],
        [/食品安全认证申请组织自我声明|自我声明/, "MSF11-27+食品安全认证申请组织自我声明（2022.1.13修订）.pdf"],
        [/监督审核通知/, "MSF11-22 监督审核通知（2015版）.pdf"],
        [/再认证通知/, "MSF11-24 再认证通知书（2015版）.pdf"],
        [/认证评定|评定表/, "MSF13-04+认证评定表（2024.3.28修订）.pdf"],
        [/认证决定|决定通知/, "MSF13-01+认证决定通知书（2015版）.pdf"],
        [/归档|材料清单|MSF11-20/, "MSF11-20+审核材料归档上报清单（2025.12.31修订）.pdf"]
    ];

    function normalizeName(value) {
        return decodeURIComponent(String(value || "归档材料预览.pdf")).trim();
    }

    function escapeHtml(value) {
        return String(value || "").replace(/[&<>"']/g, char => ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#39;"
        }[char]));
    }

    function resolvePdfFile(fileName) {
        const safeName = normalizeName(fileName);
        const exact = PDF_FILES.find(item => item === safeName || item.includes(safeName) || safeName.includes(item.replace(/\.pdf$/i, "")));
        if (exact) return exact;
        const matched = PDF_MATCHERS.find(([pattern]) => pattern.test(safeName));
        return matched ? matched[1] : "MSF11-20+审核材料归档上报清单（2025.12.31修订）.pdf";
    }

    function parseCsvLine(line) {
        const cells = [];
        let current = "";
        let quoted = false;
        for (let i = 0; i < line.length; i += 1) {
            const char = line[i];
            const next = line[i + 1];
            if (char === '"' && quoted && next === '"') {
                current += '"';
                i += 1;
            } else if (char === '"') {
                quoted = !quoted;
            } else if (char === "," && !quoted) {
                cells.push(current);
                current = "";
            } else {
                current += char;
            }
        }
        cells.push(current);
        return cells.map(cell => cell.trim());
    }

    function stripBom(value) {
        return String(value || "").replace(/^\uFEFF/, "");
    }

    function normalizeFileKey(value) {
        return stripBom(value)
            .trim()
            .replace(/\\/g, "/")
            .split("/")
            .pop()
            .replace(/\.pdf$/i, "")
            .replace(/\s+/g, "")
            .toLowerCase();
    }

    function findColumnIndex(headers, candidates, fallbackIndex) {
        const normalizedHeaders = headers.map(header => stripBom(header).trim());
        const index = normalizedHeaders.findIndex(header => candidates.includes(header));
        return index >= 0 ? index : fallbackIndex;
    }

    function registerPdfUrl(fileName, url) {
        const cleanName = stripBom(fileName).trim();
        const cleanUrl = String(url || "").trim();
        if (!cleanName || !cleanUrl) return;
        PDF_URL_MAP.set(cleanName, cleanUrl);
        PDF_URL_MAP.set(normalizeFileKey(cleanName), cleanUrl);
    }

    function loadPdfUrlMapping() {
        if (pdfUrlMappingPromise) return pdfUrlMappingPromise;
        pdfUrlMappingPromise = fetch(PDF_MAPPING_URL, { cache: "no-store" })
            .then(response => {
                if (!response.ok) throw new Error(`PDF mapping not found: ${response.status}`);
                return response.text();
            })
            .then(text => {
                let nameIndex = 1;
                let urlIndex = 2;
                text.split(/\r?\n/).forEach((line, index) => {
                    if (!line.trim()) return;
                    const cells = parseCsvLine(line);
                    if (index === 0 && stripBom(cells[0]) === "序号") {
                        nameIndex = findColumnIndex(cells, ["文件名", "PDF文件名", "pdf文件名"], 1);
                        urlIndex = findColumnIndex(cells, ["访问地址", "URL", "url", "链接", "预览地址"], 2);
                        return;
                    }
                    registerPdfUrl(cells[nameIndex], cells[urlIndex]);
                });
            })
            .catch(error => {
                console.warn("[pdf-preview] 腾讯云 PDF 映射未加载，使用项目本地 PDF 兜底：", error);
            });
        return pdfUrlMappingPromise;
    }

    function resolveRemotePdfUrl(pdfFile) {
        return PDF_URL_MAP.get(pdfFile) || PDF_URL_MAP.get(normalizeFileKey(pdfFile)) || "";
    }

    function withPdfViewerParams(url) {
        if (!url) return "";
        if (url.includes("#")) return url;
        return `${url}#toolbar=1&navpanes=0&view=FitH`;
    }

    function resolvePdfUrl(pdfFile) {
        const remoteUrl = resolveRemotePdfUrl(pdfFile);
        if (remoteUrl) {
            return {
                url: withPdfViewerParams(remoteUrl),
                fallbackUrl: withPdfViewerParams(`${PDF_BASE}${encodeURIComponent(pdfFile)}`),
                displayUrl: remoteUrl,
                sourceLabel: "腾讯云 PDF"
            };
        }
        const localUrl = `${PDF_BASE}${encodeURIComponent(pdfFile)}`;
        return {
            url: withPdfViewerParams(localUrl),
            fallbackUrl: "",
            displayUrl: `${PDF_BASE}${pdfFile}`,
            sourceLabel: "项目本地 PDF"
        };
    }

    function ensureModal() {
        let mask = document.getElementById("projectPdfPreviewMask");
        if (mask) return mask;
        mask = document.createElement("div");
        mask.id = "projectPdfPreviewMask";
        mask.className = "project-pdf-preview-mask";
        mask.innerHTML = `
            <div class="project-pdf-preview" role="dialog" aria-modal="true" aria-labelledby="projectPdfPreviewName">
                <div class="project-pdf-preview-header">
                    <div class="project-pdf-preview-title">
                        <div class="project-pdf-preview-name" id="projectPdfPreviewName">预览文件</div>
                        <div class="project-pdf-preview-meta" id="projectPdfPreviewMeta"></div>
                    </div>
                    <div class="project-pdf-preview-actions">
                        <button class="project-pdf-preview-btn" id="projectPdfPreviewFallback" type="button" onclick="useProjectPdfLocalFallback()">使用本地兜底</button>
                        <a class="project-pdf-preview-btn" id="projectPdfPreviewOpen" href="#" target="_blank" rel="noopener">新窗口打开</a>
                        <button class="project-pdf-preview-btn project-pdf-preview-close" type="button" aria-label="关闭预览" onclick="closeProjectPdfPreview()">×</button>
                    </div>
                </div>
                <div class="project-pdf-preview-body">
                    <aside class="project-pdf-preview-side">
                        <div class="project-pdf-preview-label">预览说明</div>
                        <div class="project-pdf-preview-card">
                            <div class="project-pdf-preview-card-title" id="projectPdfPreviewSourceTitle">PDF 访问链接</div>
                            <div class="project-pdf-preview-card-text" id="projectPdfPreviewPath"></div>
                        </div>
                        <div class="project-pdf-preview-card">
                            <div class="project-pdf-preview-card-title">评定口径</div>
                            <div class="project-pdf-preview-card-text">文件仅作为材料附件证据展示，问题和结论仍登记在对应 MSF11-20 材料项下。</div>
                        </div>
                    </aside>
                    <div class="project-pdf-preview-frame-wrap">
                        <iframe class="project-pdf-preview-frame" id="projectPdfPreviewFrame" title="PDF 文件预览"></iframe>
                    </div>
                </div>
            </div>
        `;
        mask.addEventListener("click", event => {
            if (event.target === mask) closeProjectPdfPreview();
        });
        document.body.appendChild(mask);
        return mask;
    }

    async function openProjectPdfPreview(fileName, options = {}) {
        await loadPdfUrlMapping();
        const businessName = normalizeName(fileName);
        const pdfFile = resolvePdfFile(options.pdfName || businessName);
        const pdfTarget = resolvePdfUrl(pdfFile);
        const mask = ensureModal();
        document.getElementById("projectPdfPreviewName").textContent = `预览：${businessName}`;
        document.getElementById("projectPdfPreviewMeta").textContent = `${pdfTarget.sourceLabel}：${pdfTarget.displayUrl}`;
        document.getElementById("projectPdfPreviewSourceTitle").textContent = pdfTarget.sourceLabel;
        document.getElementById("projectPdfPreviewPath").textContent = pdfTarget.displayUrl;
        document.getElementById("projectPdfPreviewOpen").href = pdfTarget.url;
        const fallbackBtn = document.getElementById("projectPdfPreviewFallback");
        fallbackBtn.style.display = pdfTarget.fallbackUrl ? "inline-flex" : "none";
        fallbackBtn.dataset.fallbackUrl = pdfTarget.fallbackUrl || "";
        fallbackBtn.dataset.pdfFile = pdfFile;
        const frame = document.getElementById("projectPdfPreviewFrame");
        frame.dataset.fallbackUrl = pdfTarget.fallbackUrl || "";
        frame.dataset.pdfFile = pdfFile;
        frame.src = pdfTarget.url;
        mask.classList.add("active");
    }

    function useProjectPdfLocalFallback() {
        const frame = document.getElementById("projectPdfPreviewFrame");
        if (!frame || !frame.dataset.fallbackUrl) return;
        const pdfFile = frame.dataset.pdfFile || "";
        const localDisplayUrl = `${PDF_BASE}${pdfFile}`;
        frame.src = frame.dataset.fallbackUrl;
        document.getElementById("projectPdfPreviewMeta").textContent = `项目本地 PDF：${localDisplayUrl}`;
        document.getElementById("projectPdfPreviewSourceTitle").textContent = "项目本地 PDF";
        document.getElementById("projectPdfPreviewPath").textContent = localDisplayUrl;
        document.getElementById("projectPdfPreviewOpen").href = frame.dataset.fallbackUrl;
        const fallbackBtn = document.getElementById("projectPdfPreviewFallback");
        if (fallbackBtn) fallbackBtn.style.display = "none";
    }

    function closeProjectPdfPreview() {
        const mask = document.getElementById("projectPdfPreviewMask");
        if (!mask) return;
        mask.classList.remove("active");
        const frame = document.getElementById("projectPdfPreviewFrame");
        if (frame) frame.src = "about:blank";
    }

    function renderProjectPdfInline(fileName, options = {}) {
        const businessName = normalizeName(fileName);
        const pdfFile = resolvePdfFile(options.pdfName || businessName);
        const pdfTarget = resolvePdfUrl(pdfFile);
        return `
            <div class="project-pdf-inline-note">
                当前预览载入${escapeHtml(pdfTarget.sourceLabel)}：${escapeHtml(pdfTarget.displayUrl)}。业务文件名保留为“${escapeHtml(businessName)}”，用于模拟归档系统中的附件预览。
            </div>
            <iframe class="project-pdf-inline-frame" src="${pdfTarget.url}" title="${escapeHtml(businessName)}"></iframe>
        `;
    }

    document.addEventListener("keydown", event => {
        if (event.key === "Escape") closeProjectPdfPreview();
    });

    window.PROJECT_PDF_FILES = PDF_FILES.slice();
    window.PROJECT_PDF_URL_MAP = PDF_URL_MAP;
    window.loadProjectPdfUrlMapping = loadPdfUrlMapping;
    window.resolveProjectPdfFile = resolvePdfFile;
    window.openProjectPdfPreview = openProjectPdfPreview;
    window.useProjectPdfLocalFallback = useProjectPdfLocalFallback;
    window.closeProjectPdfPreview = closeProjectPdfPreview;
    window.renderProjectPdfInline = renderProjectPdfInline;

    loadPdfUrlMapping();
}());

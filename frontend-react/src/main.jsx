import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Check,
  Download,
  ArrowDown,
  ArrowUp,
  FileText,
  FileSpreadsheet,
  FolderOpen,
  GripVertical,
  Grid2X2,
  Image,
  List,
  Loader2,
  Plus,
  RefreshCcw,
  Search,
  Shuffle,
  Trash2,
  Upload,
  X,
  Printer,
} from "lucide-react";
import "./styles.css";

const API_URL = import.meta.env.VITE_API_URL || "http://127.0.0.1:8000";
const MANY_SHEETS_LIMIT = 200;
const PAGE_SIZE = 120;
const KNOWN_BRANDS = [
  "CARAPRETA",
  "ALFAMA",
  "MINERVA",
  "GUIDARA",
  "ATIGEL",
  "FRIELLA",
  "MOCOCA",
  "PAMPLONA",
  "SADIA",
  "PERDIGAO",
];

function detectBrand(...values) {
  const text = normalizeMatchText(values.filter(Boolean).join(" "));
  return KNOWN_BRANDS.find((brand) => text.includes(normalizeMatchText(brand))) || "";
}

function productKey(item) {
  return [
    item.code || "sem-codigo",
    item.description || "",
    item.supplier || "",
    item.brand || "",
  ].join("|");
}

function productPhotoUrl(item, folderPath) {
  if (item.photoDataUrl) return item.photoDataUrl;
  if (!item.photoUrl) return "";
  const url = new URL(`${API_URL}${item.photoUrl}`);
  if (folderPath.trim()) url.searchParams.set("folder_path", folderPath.trim());
  return url.toString();
}

function apiErrorMessage(data, fallback) {
  const detail = data?.detail;
  if (!detail) return fallback;
  if (typeof detail === "string") return detail;
  if (Array.isArray(detail)) {
    return detail
      .map((item) => {
        const field = Array.isArray(item.loc) ? item.loc.join(".") : "";
        return [field, item.msg].filter(Boolean).join(": ");
      })
      .filter(Boolean)
      .join(" | ") || fallback;
  }
  if (typeof detail === "object") {
    return detail.msg || JSON.stringify(detail);
  }
  return String(detail);
}

function normalizeMatchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[^a-z0-9]+/gi, " ")
    .toLowerCase()
    .trim();
}

function matchTokens(value) {
  const ignored = new Set([
    "com",
    "das",
    "dos",
    "para",
    "pdf",
    "ftc",
    "cong",
    "congelado",
    "resfriado",
  ]);
  return normalizeMatchText(value)
    .split(/\s+/)
    .filter((token) => token.length >= 3 && !ignored.has(token) && !/^\d+$/.test(token));
}

function nameSimilarity(sheetName, pdfName) {
  const sheetText = normalizeMatchText(sheetName);
  const pdfText = normalizeMatchText(pdfName);
  if (!sheetText || !pdfText) return 0;

  const sheetCompact = sheetText.replace(/\s+/g, "");
  const pdfCompact = pdfText.replace(/\s+/g, "");
  const sheetTokens = matchTokens(sheetText);
  const pdfTokens = matchTokens(pdfText);
  let matches = 0;

  sheetTokens.forEach((sheetToken) => {
    if (
      pdfTokens.some(
        (pdfToken) =>
          pdfToken === sheetToken ||
          (pdfToken.length >= 4 && sheetToken.length >= 4 && (
            pdfToken.includes(sheetToken) ||
            sheetToken.includes(pdfToken)
          )),
      )
    ) {
      matches += 1;
    }
  });

  const tokenScore = matches ? (matches * 2) / (sheetTokens.length + pdfTokens.length) : 0;
  const compactScore =
    sheetCompact.includes(pdfCompact) || pdfCompact.includes(sheetCompact)
      ? 0.35
      : sheetTokens.some((token) => token.length >= 5 && pdfCompact.includes(token))
        ? 0.15
        : 0;

  return tokenScore + compactScore;
}

const PDF_PREVIEW_FIELDS = [
  "pdfName",
  "suggestedFile",
  "selected",
  "score",
  "brand",
  "ean",
  "boxDimensions",
  "altura",
  "largura",
  "comprimento",
  "error",
];

function App() {
  const orderedWorkbookInputRef = useRef(null);
  const orderedPdfInputRef = useRef(null);
  const orderedAddPdfInputRef = useRef(null);
  const [csvFile, setCsvFile] = useState(null);
  const [zipFile, setZipFile] = useState(null);
  const [orderedWorkbookFile, setOrderedWorkbookFile] = useState(null);
  const [orderedPdfFiles, setOrderedPdfFiles] = useState([]);
  const [supplierPdfFolderPath, setSupplierPdfFolderPath] = useState("Fichas-20260609T161612Z-3-001\\Fichas");
  const [orderedPreview, setOrderedPreview] = useState(null);
  const [orderedPreviewStale, setOrderedPreviewStale] = useState(false);
  const [draggedOrderedIndex, setDraggedOrderedIndex] = useState(null);
  const [dragOverOrderedIndex, setDragOverOrderedIndex] = useState(null);
  const [orderedLoading, setOrderedLoading] = useState(false);
  const [orderedFilling, setOrderedFilling] = useState(false);
  const [folderPath, setFolderPath] = useState("");
  const [search, setSearch] = useState("");
  const [showWithoutPhoto, setShowWithoutPhoto] = useState(false);
  const [includeSheets, setIncludeSheets] = useState(false);
  const [viewMode, setViewMode] = useState("catalog");
  const [products, setProducts] = useState([]);
  const [meta, setMeta] = useState(null);
  const [selectedItems, setSelectedItems] = useState({});
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [printingSelectedOnly, setPrintingSelectedOnly] = useState(false);
  const [error, setError] = useState("");

  const selectedCodes = useMemo(
    () => [...new Set(Object.values(selectedItems).filter(Boolean))],
    [selectedItems],
  );
  const selectedCount = selectedCodes.length;
  const isLargeSheetGeneration = includeSheets && selectedCount > MANY_SHEETS_LIMIT;
  const visibleSelectedCount = useMemo(
    () => products.filter((item) => selectedItems[productKey(item)]).length,
    [products, selectedItems],
  );
  const catalogProducts = useMemo(
    () => (printingSelectedOnly ? products.filter((item) => selectedItems[productKey(item)]) : products),
    [printingSelectedOnly, products, selectedItems],
  );

  function buildForm(extra = {}) {
    const form = new FormData();
    if (csvFile) form.append("csv_file", csvFile);
    if (zipFile) form.append("zip_file", zipFile);
    form.append("folder_path", folderPath);
    Object.entries(extra).forEach(([key, value]) => form.append(key, value));
    return form;
  }

  function buildOrderedPdfForm() {
    const form = new FormData();
    if (orderedWorkbookFile) form.append("workbook_file", orderedWorkbookFile);
    orderedPdfFiles.forEach((file) => form.append("pdf_files", file));
    return form;
  }

  function buildOrderedPdfFolderForm(extra = {}) {
    const form = new FormData();
    if (orderedWorkbookFile) form.append("workbook_file", orderedWorkbookFile);
    form.append("supplier_pdf_folder_path", supplierPdfFolderPath);
    if (orderedPreview?.items) {
      orderedPreview.items.forEach((item, index) => {
        if (item.uploadedFile) {
          form.append("replacement_pdf_files", item.uploadedFile, `replacement_${index}.pdf`);
          form.append("replacement_pdf_indexes", String(index));
        }
      });
    }
    Object.entries(extra).forEach(([key, value]) => form.append(key, value));
    return form;
  }

  function refreshFolderPreviewStats(previewData) {
    if (!previewData?.items) return previewData;
    const selectedCount = previewData.items.filter((item) => item.selected && item.suggestedFile).length;
    return {
      ...previewData,
      matchedCount: selectedCount,
      missingPdfCount: Math.max(0, (previewData.sheetCount || 0) - selectedCount),
      extraPdfCount: Math.max(0, (previewData.pdfCount || 0) - selectedCount),
    };
  }

  function reorderOrderedPdf(fromIndex, toIndex) {
    const previewItemCount = orderedPreview?.items?.length || 0;
    const reorderLimit = orderedPreview?.source === "folder" ? previewItemCount : orderedPdfFiles.length;
    if (
      !Number.isInteger(fromIndex) ||
      !Number.isInteger(toIndex) ||
      fromIndex === toIndex ||
      fromIndex < 0 ||
      toIndex < 0 ||
      fromIndex >= reorderLimit ||
      toIndex >= reorderLimit
    ) {
      return;
    }

    if (orderedPreview?.source !== "folder") {
      setOrderedPdfFiles((current) => {
        const next = [...current];
        [next[fromIndex], next[toIndex]] = [next[toIndex], next[fromIndex]];
        return next;
      });
    }
    setOrderedPreview((current) => {
      if (!current?.items) return current;
      const nextItems = current.items.map((item) => ({ ...item }));
      PDF_PREVIEW_FIELDS.forEach((field) => {
        [nextItems[fromIndex][field], nextItems[toIndex][field]] = [
          nextItems[toIndex]?.[field],
          nextItems[fromIndex]?.[field],
        ];
      });
      return {
        ...current,
        items: nextItems.map((item, itemIndex) => ({ ...item, index: itemIndex + 1 })),
      };
    });
    setOrderedPreviewStale(true);
  }

  function sortOrderedPdfsByName(previewData = orderedPreview) {
    if (!previewData?.items?.length || !orderedPdfFiles.length) return false;

    const pdfSlots = orderedPdfFiles.map((file, fileIndex) => {
      const previewItem = previewData.items[fileIndex] || {};
      return {
        file,
        fileIndex,
        pdfData: PDF_PREVIEW_FIELDS.reduce(
          (data, field) => ({
            ...data,
            [field]: previewItem[field] ?? (field === "pdfName" ? file.name : ""),
          }),
          { pdfName: previewItem.pdfName || file.name },
        ),
      };
    });
    const usedSlots = new Set();
    const selectedSlots = previewData.items.map((item) => {
      if (!item.sheetName) return null;

      let bestSlot = null;
      let bestScore = -1;
      pdfSlots.forEach((slot, slotIndex) => {
        if (usedSlots.has(slotIndex)) return;
        const score = nameSimilarity(item.sheetName, slot.pdfData.pdfName || slot.file.name);
        if (score > bestScore) {
          bestScore = score;
          bestSlot = { ...slot, slotIndex };
        }
      });

      if (bestSlot) usedSlots.add(bestSlot.slotIndex);
      return bestSlot;
    });
    const remainingSlots = pdfSlots
      .map((slot, slotIndex) => ({ ...slot, slotIndex }))
      .filter((slot) => !usedSlots.has(slot.slotIndex));
    let remainingIndex = 0;
    const orderedSlots = selectedSlots.map((slot) => {
      if (slot) return slot;
      const nextSlot = remainingSlots[remainingIndex] || null;
      remainingIndex += 1;
      return nextSlot;
    });
    remainingSlots.slice(remainingIndex).forEach((slot) => orderedSlots.push(slot));
    const nextPdfFiles = orderedSlots
      .map((slot) => slot?.file)
      .filter(Boolean);
    const orderChanged =
      nextPdfFiles.length === orderedPdfFiles.length &&
      nextPdfFiles.some((file, index) => file !== orderedPdfFiles[index]);

    if (!orderChanged) return false;

    setOrderedPdfFiles(nextPdfFiles);
    setOrderedPreview({
      ...previewData,
      items: previewData.items.map((item, itemIndex) => {
        const slot = orderedSlots[itemIndex];
        if (!slot) {
          return {
            ...item,
            index: itemIndex + 1,
            ...PDF_PREVIEW_FIELDS.reduce((data, field) => ({ ...data, [field]: "" }), {}),
          };
        }
        return {
          ...item,
          ...slot.pdfData,
          index: itemIndex + 1,
        };
      }),
    });
    setOrderedPreviewStale(true);
    return true;
  }

  function moveOrderedPdf(index, direction) {
    reorderOrderedPdf(index, index + direction);
  }

  function finishOrderedPdfDrag() {
    setDraggedOrderedIndex(null);
    setDragOverOrderedIndex(null);
  }

  function addOrderedPdfFiles(files) {
    const nextFiles = Array.from(files || []);
    if (!nextFiles.length) return;

    setOrderedPdfFiles((current) => [...current, ...nextFiles]);
    if (orderedPreview) setOrderedPreviewStale(true);
  }

  function cancelOrderedPdfFill() {
    setOrderedWorkbookFile(null);
    setOrderedPdfFiles([]);
    setSupplierPdfFolderPath("");
    setOrderedPreview(null);
    setOrderedPreviewStale(false);
    setDraggedOrderedIndex(null);
    setDragOverOrderedIndex(null);
    [orderedWorkbookInputRef, orderedPdfInputRef, orderedAddPdfInputRef].forEach((inputRef) => {
      if (inputRef.current) inputRef.current.value = "";
    });
  }

  async function loadProducts(next = {}) {
    setLoading(true);
    setError("");
    try {
      await fetch(`${API_URL}/api/health`);
      const effectiveSearch = next.search ?? search;
      const effectiveShowWithoutPhoto = next.showWithoutPhoto ?? showWithoutPhoto;
      const effectivePage = next.page ?? page;
      const response = await fetch(`${API_URL}/api/products`, {
        method: "POST",
        body: buildForm({
          search: effectiveSearch,
          only_with_photo: String(!effectiveShowWithoutPhoto),
          page: String(effectivePage),
          page_size: String(PAGE_SIZE),
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao carregar produtos."));
      setProducts(data.products);
      setMeta(data);
      setPage(data.page || effectivePage);
    } catch (err) {
      setError(
        err.message === "Failed to fetch"
          ? `Nao consegui conectar na API. Confira se o backend esta rodando em ${API_URL}.`
          : err.message,
      );
    } finally {
      setLoading(false);
    }
  }

  async function previewOrderedPdfFill() {
    if (!orderedWorkbookFile || !orderedPdfFiles.length) return;

    setOrderedLoading(true);
    setOrderedPreview(null);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/excel-pdf-order/preview`, {
        method: "POST",
        body: buildOrderedPdfForm(),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao conferir ordem dos PDFs."));
      data.source = "upload";
      const sortedByName = sortOrderedPdfsByName(data);
      if (!sortedByName) setOrderedPreview(data);
      setOrderedPreviewStale(sortedByName);
    } catch (err) {
      setError(err.message);
    } finally {
      setOrderedLoading(false);
    }
  }

  async function suggestOrderedPdfsFromFolder() {
    if (!orderedWorkbookFile || !supplierPdfFolderPath.trim()) return;

    setOrderedLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/excel-pdf-order/suggest-folder`, {
        method: "POST",
        body: buildOrderedPdfFolderForm(),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao sugerir fichas."));
      setOrderedPreview(refreshFolderPreviewStats({ ...data, source: "folder" }));
      setOrderedPreviewStale(true);
    } catch (err) {
      setError(err.message);
    } finally {
      setOrderedLoading(false);
    }
  }

  function toggleSuggestedPdf(index) {
    setOrderedPreview((current) => {
      if (!current?.items) return current;
      const next = {
        ...current,
        items: current.items.map((item, itemIndex) =>
          itemIndex === index ? { ...item, selected: !item.selected } : item,
        ),
      };
      return current.source === "folder" ? refreshFolderPreviewStats(next) : next;
    });
    setOrderedPreviewStale(true);
  }

  async function replaceSuggestedPdfWithUpload(index, file) {
    if (!file) return;

    setError("");
    try {
      const form = new FormData();
      form.append("supplier_pdf_file", file);

      const response = await fetch(`${API_URL}/api/supplier-pdf/extract`, {
        method: "POST",
        body: form,
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao ler ficha escolhida."));

      setOrderedPreview((current) => {
        if (!current?.items) return current;
        const next = {
          ...current,
          items: current.items.map((item, itemIndex) =>
            itemIndex === index
              ? {
                  ...item,
                  pdfName: file.name,
                  suggestedFile: "",
                  uploadedFile: file,
                  selected: true,
                  score: 100,
                  brand: detectBrand(file.name, item.productDescription, item.sheetName),
                  ean: data.ean || "",
                  boxDimensions: data.box_dimensions || "",
                  altura: data.altura || "",
                  largura: data.largura || "",
                  comprimento: data.comprimento || "",
                  error: "",
                }
              : item,
          ),
        };
        return current.source === "folder" ? refreshFolderPreviewStats(next) : next;
      });
      setOrderedPreviewStale(true);
    } catch (err) {
      setError(err.message);
    }
  }

  async function fillWorkbookByPdfOrder() {
    if (
      !orderedWorkbookFile ||
      (orderedPreview?.source === "folder"
        ? !orderedPreview.items?.some((item) => item.selected && item.suggestedFile)
        : !orderedPdfFiles.length)
    ) return;

    setOrderedFilling(true);
    setError("");
    try {
      const isFolderPreview = orderedPreview?.source === "folder";
      const response = await fetch(`${API_URL}/api/excel-pdf-order/${isFolderPreview ? "fill-folder" : "fill"}`, {
        method: "POST",
        body: isFolderPreview
          ? buildOrderedPdfFolderForm({
              selected_pdf_files: JSON.stringify(
                orderedPreview.items.map((item) =>
                  item.selected && !item.uploadedFile ? item.suggestedFile || item.pdfName || "" : "",
                ),
              ),
            })
          : buildOrderedPdfForm(),
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(apiErrorMessage(data, "Falha ao preencher Excel com PDFs."));
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "produtos_com_fotos_preenchido.xlsx";
      link.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setOrderedFilling(false);
    }
  }

  useEffect(() => {
    loadProducts();
  }, []);

  function toggleProduct(item) {
    if (!item.code) return;
    const key = productKey(item);
    setSelectedItems((current) => {
      const next = { ...current };
      if (next[key]) delete next[key];
      else next[key] = item.code;
      return next;
    });
  }

  function selectVisibleWithPhoto() {
    setSelectedItems((current) => {
      const next = { ...current };
      products.forEach((item) => {
        if (item.hasPhoto && item.code) next[productKey(item)] = item.code;
      });
      return next;
    });
  }

  async function generateWorkbook() {
    if (!selectedCount) return;
    if (isLargeSheetGeneration) {
      setError(
        `Gerar ${selectedCount} fichas individuais pode demorar muito. Desmarque "Fichas individuais" para gerar a planilha rapidamente.`,
      );
      return;
    }

    setGenerating(true);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/generate`, {
        method: "POST",
        body: buildForm({
          selected_codes: JSON.stringify([...selectedCodes]),
          include_product_sheets: String(includeSheets),
        }),
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(apiErrorMessage(data, "Falha ao gerar planilha."));
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "produtos_com_fotos.xlsx";
      link.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setGenerating(false);
    }
  }

  function exportCatalogPdf() {
    if (viewMode !== "catalog") setViewMode("catalog");
    setPrintingSelectedOnly(selectedCount > 0);
    window.setTimeout(() => {
      window.print();
      window.setTimeout(() => setPrintingSelectedOnly(false), 250);
    }, 120);
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <h1>Catalogo de produtos</h1>
          <p>Veja os produtos com fotos, selecione itens e gere a planilha Excel.</p>
        </div>
        <div className="top-actions">
          <div className="segmented" aria-label="Modo de visualizacao">
            <button
              className={viewMode === "catalog" ? "active" : ""}
              onClick={() => setViewMode("catalog")}
              title="Catalogo"
            >
              <Grid2X2 size={17} />
            </button>
            <button
              className={viewMode === "table" ? "active" : ""}
              onClick={() => setViewMode("table")}
              title="Tabela"
            >
              <List size={17} />
            </button>
          </div>
          <button className="primary" onClick={generateWorkbook} disabled={!selectedCount || generating}>
            {generating ? <Loader2 className="spin" size={18} /> : <Download size={18} />}
            {includeSheets ? "Gerar Excel com fichas" : "Gerar Excel"}
          </button>
          <button className="secondary" onClick={exportCatalogPdf} disabled={!products.length}>
            <Printer size={18} />
            {selectedCount ? "Exportar selecionados" : "Exportar PDF"}
          </button>
        </div>
      </header>

      <section className="control-band">
        <label className="file-control">
          <Upload size={17} />
          <span>{csvFile ? csvFile.name : "CSV de produtos"}</span>
          <input type="file" accept=".csv" onChange={(event) => setCsvFile(event.target.files?.[0] || null)} />
        </label>

        <label className="file-control">
          <Image size={17} />
          <span>{zipFile ? zipFile.name : "ZIP das fotos"}</span>
          <input type="file" accept=".zip" onChange={(event) => setZipFile(event.target.files?.[0] || null)} />
        </label>

        <label className="folder-input">
          <FolderOpen size={17} />
          <input
            value={folderPath}
            onChange={(event) => setFolderPath(event.target.value)}
            placeholder="Caminho da pasta de fotos"
          />
        </label>

        <button className="secondary" onClick={() => loadProducts({ page: 1 })} disabled={loading}>
          {loading ? <Loader2 className="spin" size={17} /> : <RefreshCcw size={17} />}
          Carregar
        </button>
      </section>

      <section className="filter-band">
        <label className="search-input">
          <Search size={17} />
          <input
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") loadProducts({ search: event.currentTarget.value, page: 1 });
            }}
            placeholder="Buscar por codigo, descricao, fornecedor ou marca"
          />
        </label>

        <label className="check-control">
          <input
            type="checkbox"
            checked={showWithoutPhoto}
            onChange={(event) => {
              setShowWithoutPhoto(event.target.checked);
              loadProducts({ showWithoutPhoto: event.target.checked, page: 1 });
            }}
          />
          Mostrar sem foto
        </label>

        <label className="check-control">
          <input
            type="checkbox"
            checked={includeSheets}
            onChange={(event) => setIncludeSheets(event.target.checked)}
          />
          Fichas individuais
        </label>

        <button className="ghost" onClick={selectVisibleWithPhoto}>
          <Check size={17} />
          Selecionar com foto
        </button>

        <button className="ghost danger" onClick={() => setSelectedItems({})}>
          <Trash2 size={17} />
          Limpar selecao
        </button>
      </section>

      <section className="ordered-pdf-tool">
        <div className="ordered-pdf-controls">
          <label className="file-control">
            <FileSpreadsheet size={17} />
            <span>{orderedWorkbookFile ? orderedWorkbookFile.name : "Excel gerado"}</span>
            <input
              ref={orderedWorkbookInputRef}
              type="file"
              accept=".xlsx"
              onChange={(event) => {
                setOrderedWorkbookFile(event.target.files?.[0] || null);
                setOrderedPreview(null);
                event.target.value = "";
              }}
            />
          </label>

          <label className="file-control">
            <FileText size={17} />
            <span>
              {orderedPdfFiles.length ? `${orderedPdfFiles.length} PDFs em ordem` : "PDFs na ordem das abas"}
            </span>
            <input
              ref={orderedPdfInputRef}
              type="file"
              accept=".pdf"
              multiple
              onChange={(event) => {
                setOrderedPdfFiles(Array.from(event.target.files || []));
                setOrderedPreview(null);
                setOrderedPreviewStale(false);
                event.target.value = "";
              }}
            />
          </label>

          <label className="file-control add-file-control">
            <Plus size={17} />
            <span>Adicionar PDFs</span>
            <input
              ref={orderedAddPdfInputRef}
              type="file"
              accept=".pdf"
              multiple
              onChange={(event) => {
                addOrderedPdfFiles(event.target.files);
                event.target.value = "";
              }}
            />
          </label>

          <label className="folder-input pdf-folder-input">
            <FolderOpen size={17} />
            <input
              value={supplierPdfFolderPath}
              onChange={(event) => setSupplierPdfFolderPath(event.target.value)}
              placeholder="Pasta das fichas PDF"
            />
          </label>

          <button
            className="ghost"
            onClick={suggestOrderedPdfsFromFolder}
            disabled={orderedLoading || !orderedWorkbookFile || !supplierPdfFolderPath.trim()}
          >
            {orderedLoading ? <Loader2 className="spin" size={17} /> : <Search size={17} />}
            Sugerir fichas
          </button>

          <button
            className="secondary"
            onClick={previewOrderedPdfFill}
            disabled={orderedLoading || !orderedWorkbookFile || !orderedPdfFiles.length}
          >
            {orderedLoading ? <Loader2 className="spin" size={17} /> : <Search size={17} />}
            Conferir ordem
          </button>

          <button
            className="ghost"
            onClick={() => sortOrderedPdfsByName()}
            disabled={!orderedPreview || orderedLoading || !orderedPdfFiles.length}
          >
            <Shuffle size={17} />
            Organizar por nome
          </button>

          <button
            className="primary"
            onClick={fillWorkbookByPdfOrder}
            disabled={
              orderedFilling ||
              !orderedWorkbookFile ||
              (orderedPreview?.source === "folder"
                ? !orderedPreview.items?.some((item) => item.selected && item.suggestedFile)
                : !orderedPdfFiles.length)
            }
          >
            {orderedFilling ? <Loader2 className="spin" size={17} /> : <Download size={17} />}
            Baixar Excel preenchido
          </button>

          <button
            className="ghost danger"
            onClick={cancelOrderedPdfFill}
            disabled={orderedLoading || orderedFilling || (!orderedWorkbookFile && !orderedPdfFiles.length && !supplierPdfFolderPath && !orderedPreview)}
          >
            <X size={17} />
            Cancelar
          </button>
        </div>

        {orderedPreview && (
          <div className="ordered-preview">
            <div className="pdf-analysis-summary">
              <span>Abas: {orderedPreview.sheetCount}</span>
              <span>PDFs: {orderedPreview.pdfCount}</span>
              <span className="ok">Pareados: {orderedPreview.matchedCount}</span>
              <span className={orderedPreview.missingPdfCount ? "warn" : "ok"}>
                Sem PDF: {orderedPreview.missingPdfCount}
              </span>
              <span className={orderedPreview.extraPdfCount ? "warn" : "ok"}>
                PDFs extras: {orderedPreview.extraPdfCount}
              </span>
              {orderedPreviewStale && <span className="warn">Ordem alterada: conferir novamente</span>}
            </div>

            <div className="pdf-match-list">
              <div className="pdf-match-header ordered-row">
                <span>Acoes</span>
                <span>Ficha no Excel</span>
                <span>PDF usado</span>
                <span>Marca</span>
                <span>EAN</span>
                <span>Dimensoes</span>
              </div>
              {orderedPreview.items.slice(0, 20).map((item) => {
                const itemIndex = item.index - 1;
                const moveLimit = orderedPreview.source === "folder"
                  ? orderedPreview.items.length
                  : orderedPdfFiles.length;
                return (
                <div
                  className={[
                    item.error ? "pdf-match-row ordered-row" : "pdf-match-row ordered-row matched",
                    draggedOrderedIndex === itemIndex ? "dragging" : "",
                    dragOverOrderedIndex === itemIndex && draggedOrderedIndex !== itemIndex ? "drag-over" : "",
                  ].filter(Boolean).join(" ")}
                  key={`${item.index}-${item.pdfName}`}
                  draggable
                  onDragStart={(event) => {
                    setDraggedOrderedIndex(itemIndex);
                    event.dataTransfer.effectAllowed = "move";
                    event.dataTransfer.setData("text/plain", String(itemIndex));
                  }}
                  onDragEnter={(event) => {
                    event.preventDefault();
                    setDragOverOrderedIndex(itemIndex);
                  }}
                  onDragOver={(event) => {
                    event.preventDefault();
                    event.dataTransfer.dropEffect = "move";
                  }}
                  onDragLeave={(event) => {
                    if (!event.relatedTarget || !event.currentTarget.contains(event.relatedTarget)) {
                      setDragOverOrderedIndex((current) => (current === itemIndex ? null : current));
                    }
                  }}
                  onDrop={(event) => {
                    event.preventDefault();
                    const fromIndex = Number(event.dataTransfer.getData("text/plain"));
                    reorderOrderedPdf(Number.isNaN(fromIndex) ? draggedOrderedIndex : fromIndex, itemIndex);
                    finishOrderedPdfDrag();
                  }}
                  onDragEnd={finishOrderedPdfDrag}
                >
                  <span className="row-actions">
                    {orderedPreview.source === "folder" && (
                      <input
                        type="checkbox"
                        checked={Boolean(item.selected)}
                        disabled={!item.suggestedFile}
                        onChange={() => toggleSuggestedPdf(itemIndex)}
                        title={item.selected ? "Usar esta ficha" : "Ignorar esta ficha"}
                        aria-label={item.selected ? "Usar esta ficha" : "Ignorar esta ficha"}
                      />
                    )}
                    <span className="drag-handle" title="Arrastar PDF para outra posicao">
                      <GripVertical size={15} />
                    </span>
                    <button
                      className="icon-button"
                      onClick={() => moveOrderedPdf(itemIndex, -1)}
                      disabled={item.index <= 1}
                      title="Subir PDF"
                    >
                      <ArrowUp size={15} />
                    </button>
                    <button
                      className="icon-button"
                      onClick={() => moveOrderedPdf(itemIndex, 1)}
                      disabled={item.index >= moveLimit}
                      title="Descer PDF"
                    >
                      <ArrowDown size={15} />
                    </button>
                  </span>
                  <span>{item.index}. {item.sheetName || "Sem aba"}</span>
                  <span className="pdf-file-cell" title={item.pdfName}>
                    <span>
                      {orderedPreview.source === "folder" && !item.selected && item.pdfName
                        ? `Ignorada: ${item.pdfName}`
                        : item.pdfName}
                    </span>
                    {orderedPreview.source === "folder" && (
                      <label className="replace-button">
                        Substituir
                        <input
                          type="file"
                          accept=".pdf"
                          onChange={(event) => {
                            replaceSuggestedPdfWithUpload(itemIndex, event.target.files?.[0] || null);
                            event.target.value = "";
                          }}
                        />
                      </label>
                    )}
                  </span>
                  <span>{item.brand || "Marca nao encontrada"}</span>
                  <span>{item.ean || item.error || "EAN nao encontrado"}</span>
                  <span>{item.boxDimensions || "Sem dimensoes"}</span>
                </div>
              );
              })}
            </div>
          </div>
        )}
      </section>

      {error && <div className="alert">{error}</div>}

      {isLargeSheetGeneration && (
        <div className="alert">
          Muitos produtos selecionados para fichas individuais. Para gerar todos os produtos com velocidade,
          desmarque "Fichas individuais" e mantenha apenas a aba Produtos.
        </div>
      )}

      <section className="status-line">
        <span><FileSpreadsheet size={16} /> Produtos: {meta?.total ?? 0}</span>
        <span>Filtrados: {meta?.filteredTotal ?? 0}</span>
        <span>Pagina: {meta?.page ?? page}</span>
        <span>Fotos indexadas: {meta?.photoCount ?? 0}</span>
        <span>Selecionados: {selectedCount}</span>
        <span>Visiveis selecionados: {visibleSelectedCount}</span>
        <span className={meta?.templateFound ? "ok" : "warn"}>
          Modelo: {meta?.templateFound ? meta.templateName : "nao encontrado"}
        </span>
      </section>

      <section className="pagination-band">
        <button
          className="secondary"
          onClick={() => loadProducts({ page: Math.max(1, page - 1) })}
          disabled={loading || page <= 1}
        >
          Anterior
        </button>
        <span>
          {products.length ? `${(page - 1) * PAGE_SIZE + 1}-${(page - 1) * PAGE_SIZE + products.length}` : "0"} de {meta?.filteredTotal ?? 0}
        </span>
        <button
          className="secondary"
          onClick={() => loadProducts({ page: page + 1 })}
          disabled={loading || !meta?.hasNextPage}
        >
          Proxima
        </button>
      </section>

      {viewMode === "catalog" && (
        <section className="catalog-grid">
          {catalogProducts.map((item) => {
            const photoUrl = productPhotoUrl(item, folderPath);
            const key = productKey(item);
            return (
              <article
                className={`product-card ${selectedItems[key] ? "selected" : ""} ${!item.code ? "disabled" : ""}`}
                key={key}
                onClick={() => toggleProduct(item)}
              >
                <div className="product-image">
                  {photoUrl ? (
                    <img src={photoUrl} alt={item.description || `Produto ${item.code}`} loading="lazy" />
                  ) : (
                    <Image size={34} />
                  )}
                  <input
                    type="checkbox"
                    checked={Boolean(selectedItems[key])}
                    disabled={!item.code}
                    onChange={() => toggleProduct(item)}
                    onClick={(event) => event.stopPropagation()}
                    aria-label={`Selecionar ${item.description || item.code}`}
                  />
                </div>
                <div className="product-info">
                  <span className="product-code">{item.code || "Sem codigo"}</span>
                  <h2>{item.description || "Produto sem descricao"}</h2>
                  <p>{item.brand || item.supplier || "Sem marca informada"}</p>
                  <span className={item.hasPhoto ? "photo yes" : "photo no"}>
                    {item.hasPhoto ? "COM FOTO" : "SEM FOTO"}
                  </span>
                </div>
              </article>
            );
          })}
          {!loading && catalogProducts.length === 0 && (
            <div className="empty-panel">
              {printingSelectedOnly ? "Nenhum produto selecionado nesta pagina." : "Nenhum produto carregado."}
            </div>
          )}
          {loading && <div className="empty-panel">Carregando produtos...</div>}
        </section>
      )}

      {viewMode === "table" && (
        <section className="table-wrap">
          <table>
            <thead>
              <tr>
                <th className="select-col"></th>
                <th>Codigo</th>
                <th>Descricao</th>
                <th>Fornecedor</th>
                <th>Marca</th>
                <th>Foto</th>
              </tr>
            </thead>
            <tbody>
              {products.map((item) => {
                const key = productKey(item);
                return (
                <tr key={key}>
                  <td>
                    <input
                      type="checkbox"
                      checked={Boolean(selectedItems[key])}
                      disabled={!item.code}
                      onChange={() => toggleProduct(item)}
                    />
                  </td>
                  <td className="code-cell">{item.code}</td>
                  <td>{item.description}</td>
                  <td>{item.supplier}</td>
                  <td>{item.brand}</td>
                  <td>
                    <span className={item.hasPhoto ? "photo yes" : "photo no"}>
                      {item.hasPhoto ? "SIM" : "NAO"}
                    </span>
                  </td>
                </tr>
              );
              })}
              {!loading && products.length === 0 && (
                <tr>
                  <td colSpan="6" className="empty-row">Nenhum produto carregado.</td>
                </tr>
              )}
              {loading && (
                <tr>
                  <td colSpan="6" className="empty-row">Carregando produtos...</td>
                </tr>
              )}
            </tbody>
          </table>
        </section>
      )}
    </main>
  );
}

createRoot(document.getElementById("root")).render(<App />);

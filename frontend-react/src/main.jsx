import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Check,
  Download,
  ArrowDown,
  ArrowUp,
  FileText,
  FileCode2,
  FileSpreadsheet,
  FolderOpen,
  GripVertical,
  Grid2X2,
  Image,
  List,
  Loader2,
  LogOut,
  MoreHorizontal,
  Pencil,
  Plus,
  RefreshCcw,
  Search,
  Shuffle,
  Trash2,
  Upload,
  UserPlus,
  Users,
  X,
  Printer,
  DollarSign,
  User,
} from "lucide-react";
import "./styles.css";

function resolveApiUrl() {
  const configuredUrl = String(import.meta.env.VITE_API_URL || "").trim();
  const browserHost = window.location.hostname || "127.0.0.1";
  if (configuredUrl) {
    return configuredUrl.replace(
      /^(https?:\/\/)(localhost|127\.0\.0\.1)(:\d+)?/i,
      (_, protocol, _localHost, port = "") => `${protocol}${browserHost}${port}`,
    );
  }
  return `${window.location.protocol}//${browserHost}:8000`;
}

const API_URL = resolveApiUrl();
const MANY_SHEETS_LIMIT = 200;
const PAGE_SIZE = 120;
const DEFAULT_SUPPLIER_PDF_FOLDER = "Fichas-20260609T161612Z-3-001\\Fichas";
const KNOWN_BRANDS = [
  "CARAPRETA",
  "ALFAMA",
  "MINERVA",
  "COOPAVEL",
  "VALENCIO",
  "GUIDARA",
  "ATIGEL",
  "AVE NOVA",
  "DUBOI",
  "DAUS",
  "EASYCHEF",
  "FRIELLA",
  "MARIZA",
  "MOCOCA",
  "PAMPLONA",
  "PLENA",
  "TUDBOM",
  "SADIA",
  "PERDIGAO",
  "RIO MARIA",
  "RAINHA DA PAZ",
  "SOMAVE",
  "STELLADORO",
  "SAO FRANCISCO",
  "SAO VICENTE",
  "TEMPERO DA CASA",
];

function storedSession() {
  try {
    const session = JSON.parse(localStorage.getItem("productSystemSession") || "null");
    return session?.token && session?.user?.username && session?.user?.role ? session : null;
  } catch {
    localStorage.removeItem("productSystemSession");
    return null;
  }
}

function formatPrice(value) {
  const number = Number(value);
  return Number.isFinite(number)
    ? number.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })
    : "Sem preço";
}

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function submitLogin(event) {
    event.preventDefault();
    setLoading(true);
    setError("");
    try {
      const form = new FormData();
      form.append("username", username);
      form.append("password", password);
      const response = await fetch(`${API_URL}/api/login`, { method: "POST", body: form });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao entrar."));
      onLogin(data);
    } catch (err) {
      setError(
        err.message === "Failed to fetch"
          ? `Nao consegui conectar na API em ${API_URL}.`
          : err.message,
      );
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login-page">
      <form className="login-card" onSubmit={submitLogin}>
        <div className="login-brand">
          <FileSpreadsheet size={28} />
          <div>
            <h1>Sistema de produtos</h1>
            <p>Entre com seu usuário e senha</p>
          </div>
        </div>
        {error && <div className="alert">{error}</div>}
        <label>
          Usuário
          <input
            autoComplete="username"
            autoFocus
            value={username}
            onChange={(event) => setUsername(event.target.value)}
            placeholder="seu usuário"
          />
        </label>
        <label>
          Senha
          <input
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        <button className="primary" type="submit" disabled={loading || !username || !password}>
          {loading ? <Loader2 className="spin" size={18} /> : <User size={18} />}
          Entrar
        </button>
      </form>
    </main>
  );
}

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
  const xmlInputRef = useRef(null);
  const [csvFile, setCsvFile] = useState(null);
  const [zipFile, setZipFile] = useState(null);
  const [orderedWorkbookFile, setOrderedWorkbookFile] = useState(null);
  const [orderedPdfFiles, setOrderedPdfFiles] = useState([]);
  const [supplierPdfFolderPath, setSupplierPdfFolderPath] = useState(DEFAULT_SUPPLIER_PDF_FOLDER);
  const [showPdfFolderEditor, setShowPdfFolderEditor] = useState(false);
  const [orderedPreview, setOrderedPreview] = useState(null);
  const [orderedPreviewStale, setOrderedPreviewStale] = useState(false);
  const [draggedOrderedIndex, setDraggedOrderedIndex] = useState(null);
  const [dragOverOrderedIndex, setDragOverOrderedIndex] = useState(null);
  const [orderedLoading, setOrderedLoading] = useState(false);
  const [orderedFilling, setOrderedFilling] = useState(false);
  const [xmlFiles, setXmlFiles] = useState([]);
  const [xmlPreview, setXmlPreview] = useState(null);
  const [xmlLoading, setXmlLoading] = useState(false);
  const [xmlFilling, setXmlFilling] = useState(false);
  const [folderPath, setFolderPath] = useState("");
  const [search, setSearch] = useState("");
  const [showWithoutPhoto, setShowWithoutPhoto] = useState(false);
  const [includeSheets, setIncludeSheets] = useState(false);
  const [includePrices, setIncludePrices] = useState(false);
  const [pricePanelOpen, setPricePanelOpen] = useState(false);
  const [viewMode, setViewMode] = useState("catalog");
  const [products, setProducts] = useState([]);
  const [meta, setMeta] = useState(null);
  const [selectedItems, setSelectedItems] = useState({});
  const [selectedProducts, setSelectedProducts] = useState({});
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [printingSelectedOnly, setPrintingSelectedOnly] = useState(false);
  const [error, setError] = useState("");
  const [activeSection, setActiveSection] = useState("generator");
  const [session, setSession] = useState(storedSession);
  const [editedPrices, setEditedPrices] = useState({});
  const [users, setUsers] = useState([]);
  const [usersLoading, setUsersLoading] = useState(false);
  const [pdfAudit, setPdfAudit] = useState(null);
  const [pdfAuditLoading, setPdfAuditLoading] = useState(false);
  const [pdfAuditSearch, setPdfAuditSearch] = useState("");
  const [pdfAuditOnlyCompatible, setPdfAuditOnlyCompatible] = useState(false);
  const [pdfAuditMessage, setPdfAuditMessage] = useState("");
  const [editingUserId, setEditingUserId] = useState(null);
  const [userForm, setUserForm] = useState({
    username: "",
    password: "",
    role: "vendedor",
    active: true,
  });

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
  const selectedPriceProducts = useMemo(
    () => selectedCodes.map((code) => selectedProducts[code]).filter(Boolean),
    [selectedCodes, selectedProducts],
  );
  const filteredPdfAuditItems = useMemo(() => {
    const query = pdfAuditSearch.trim().toLowerCase();
    return (pdfAudit?.items || [])
      .map((item, auditIndex) => ({ ...item, auditIndex }))
      .filter((item) => !pdfAuditOnlyCompatible || Number(item.score || 0) > 50)
      .filter((item) =>
        !query || [item.productCode, item.productDescription, item.supplier, item.factoryCode, item.suggestedFile]
          .join(" ")
          .toLowerCase()
          .includes(query),
      );
  }, [pdfAudit, pdfAuditOnlyCompatible, pdfAuditSearch]);

  function buildForm(extra = {}) {
    const form = new FormData();
    if (csvFile) form.append("csv_file", csvFile);
    if (zipFile) form.append("zip_file", zipFile);
    form.append("folder_path", folderPath);
    Object.entries(extra).forEach(([key, value]) => form.append(key, value));
    return form;
  }

  function authHeaders(token = session?.token) {
    return token ? { Authorization: `Bearer ${token}` } : {};
  }

  function saveSession(nextSession) {
    localStorage.setItem("productSystemSession", JSON.stringify(nextSession));
    setSession(nextSession);
  }

  function logout() {
    localStorage.removeItem("productSystemSession");
    setSession(null);
    setProducts([]);
    setSelectedItems({});
    setSelectedProducts({});
    setEditedPrices({});
    setError("");
  }

  function resetUserForm() {
    setEditingUserId(null);
    setUserForm({ username: "", password: "", role: "vendedor", active: true });
  }

  async function loadUsers(token) {
    setUsersLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/users`, { headers: authHeaders(token) });
      const data = await response.json();
      if (response.status === 401) logout();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao carregar usuários."));
      setUsers(data.users || []);
    } catch (err) {
      setError(err.message);
    } finally {
      setUsersLoading(false);
    }
  }

  async function saveUser(event) {
    event.preventDefault();
    setUsersLoading(true);
    setError("");
    try {
      const payload = {
        username: userForm.username,
        role: userForm.role,
      };
      if (editingUserId) payload.active = userForm.active;
      if (userForm.password) payload.password = userForm.password;
      const response = await fetch(
        editingUserId ? `${API_URL}/api/users/${editingUserId}` : `${API_URL}/api/users`,
        {
          method: editingUserId ? "PUT" : "POST",
          headers: { ...authHeaders(), "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        },
      );
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao salvar usuário."));
      if (data.token) {
        saveSession({
          token: data.token,
          user: { id: data.id, username: data.username, role: data.role },
        });
      }
      resetUserForm();
      await loadUsers(data.token);
    } catch (err) {
      setError(err.message);
      setUsersLoading(false);
    }
  }

  function editUser(user) {
    if (user.fixed) return;
    setEditingUserId(user.id);
    setUserForm({
      username: user.username,
      password: "",
      role: user.role,
      active: user.active,
    });
  }

  async function removeUser(user) {
    if (!window.confirm(`Excluir o usuário ${user.username}?`)) return;
    setUsersLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/users/${user.id}`, {
        method: "DELETE",
        headers: authHeaders(),
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(apiErrorMessage(data, "Falha ao excluir usuário."));
      }
      if (editingUserId === user.id) resetUserForm();
      await loadUsers();
    } catch (err) {
      setError(err.message);
      setUsersLoading(false);
    }
  }

  async function loadPdfAudit() {
    setPdfAuditLoading(true);
    setPdfAuditMessage("");
    setError("");
    try {
      const form = new FormData();
      form.append("supplier_pdf_folder_path", supplierPdfFolderPath);
      form.append("folder_path", folderPath);
      const response = await fetch(`${API_URL}/api/admin/pdf-audit/suggest`, {
        method: "POST",
        headers: authHeaders(),
        body: form,
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao auditar as fichas."));
      setPdfAudit(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setPdfAuditLoading(false);
    }
  }

  function updatePdfAuditItem(index, changes) {
    setPdfAudit((current) => ({
      ...current,
      items: current.items.map((item, itemIndex) =>
        itemIndex === index ? { ...item, ...changes } : item,
      ),
    }));
  }

  async function renameAuditedPdfs() {
    const items = (pdfAudit?.items || [])
      .filter((item) => item.selected && item.suggestedFile)
      .map((item) => ({ productCode: item.productCode, sourceFile: item.suggestedFile }));
    if (!items.length) return;
    if (!window.confirm(`Renomear ${items.length} ficha(s) adicionando o codigo interno?`)) return;

    setPdfAuditLoading(true);
    setPdfAuditMessage("");
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/admin/pdf-audit/rename`, {
        method: "POST",
        headers: { ...authHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify({ folderPath: pdfAudit.folderPath, items }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao renomear as fichas."));
      setPdfAuditMessage(`${data.renamedCount} ficha(s) renomeada(s) com sucesso.`);
      await loadPdfAudit();
      setPdfAuditMessage(`${data.renamedCount} ficha(s) renomeada(s) com sucesso.`);
    } catch (err) {
      setError(err.message);
    } finally {
      setPdfAuditLoading(false);
    }
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

  function buildXmlCestForm(extra = {}) {
    const form = new FormData();
    if (orderedWorkbookFile) form.append("workbook_file", orderedWorkbookFile);
    xmlFiles.forEach((file) => form.append("xml_files", file));
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
    setXmlFiles([]);
    setXmlPreview(null);
    [orderedWorkbookInputRef, orderedPdfInputRef, orderedAddPdfInputRef, xmlInputRef].forEach((inputRef) => {
      if (inputRef.current) inputRef.current.value = "";
    });
  }

  async function previewXmlCest() {
    if (!orderedWorkbookFile || !xmlFiles.length) return;
    setXmlLoading(true);
    setXmlPreview(null);
    setError("");
    try {
      const response = await fetch(`${API_URL}/api/excel-xml-cest/preview`, {
        method: "POST",
        headers: authHeaders(),
        body: buildXmlCestForm(),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(apiErrorMessage(data, "Falha ao conferir CEST dos XMLs."));
      setXmlPreview(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setXmlLoading(false);
    }
  }

  function toggleXmlCestItem(index) {
    setXmlPreview((current) => ({
      ...current,
      items: current.items.map((item, itemIndex) =>
        itemIndex === index ? { ...item, selected: !item.selected } : item,
      ),
    }));
  }

  async function fillWorkbookWithXmlCest() {
    if (!orderedWorkbookFile || !xmlFiles.length || !xmlPreview) return;
    setXmlFilling(true);
    setError("");
    try {
      const selectedIndexes = xmlPreview.items
        .map((item, index) => (item.selected ? index : null))
        .filter((index) => index !== null);
      const response = await fetch(`${API_URL}/api/excel-xml-cest/fill`, {
        method: "POST",
        headers: authHeaders(),
        body: buildXmlCestForm({ selected_indexes: JSON.stringify(selectedIndexes) }),
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(apiErrorMessage(data, "Falha ao preencher CEST no Excel."));
      }
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "produtos_com_cest.xlsx";
      link.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setXmlFilling(false);
    }
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
        headers: authHeaders(),
        body: buildForm({
          search: effectiveSearch,
          only_with_photo: String(!effectiveShowWithoutPhoto),
          page: String(effectivePage),
          page_size: String(PAGE_SIZE),
        }),
      });
      const data = await response.json();
      if (response.status === 401) logout();
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
    if (session?.token) loadProducts();
  }, [session?.token]);

  function toggleProduct(item) {
    if (!item.code) return;
    const key = productKey(item);
    const wasSelected = Boolean(selectedItems[key]);
    setSelectedItems((current) => {
      const next = { ...current };
      if (next[key]) delete next[key];
      else next[key] = item.code;
      return next;
    });
    setSelectedProducts((current) => {
      const next = { ...current };
      if (wasSelected) delete next[item.code];
      else next[item.code] = item;
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
    setSelectedProducts((current) => {
      const next = { ...current };
      products.forEach((item) => {
        if (item.hasPhoto && item.code) next[item.code] = item;
      });
      return next;
    });
  }

  function clearSelection() {
    setSelectedItems({});
    setSelectedProducts({});
    setEditedPrices({});
  }

  function removeSelectedProduct(item) {
    setSelectedItems((current) => {
      const next = { ...current };
      delete next[productKey(item)];
      return next;
    });
    setSelectedProducts((current) => {
      const next = { ...current };
      delete next[item.code];
      return next;
    });
    setEditedPrices((current) => {
      const next = { ...current };
      delete next[item.code];
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
      const selectedCustomPrices = Object.fromEntries(
        selectedCodes
          .filter((code) => editedPrices[code] !== undefined && editedPrices[code] !== "")
          .map((code) => [code, editedPrices[code]]),
      );
      const response = await fetch(`${API_URL}/api/generate`, {
        method: "POST",
        headers: authHeaders(),
        body: buildForm({
          selected_codes: JSON.stringify([...selectedCodes]),
          include_product_sheets: String(includeSheets),
          include_prices: String(includePrices),
          custom_prices: JSON.stringify(includePrices ? selectedCustomPrices : {}),
          supplier_pdf_folder_path: includeSheets ? supplierPdfFolderPath : "",
        }),
      });
      if (!response.ok) {
        const data = await response.json();
        if (response.status === 401) logout();
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

  if (!session) return <LoginScreen onLogin={saveSession} />;

  return (
    <div className="app-layout">
      <aside className="sidebar">
        <div className="sidebar-brand">
          <FileSpreadsheet size={22} />
          <div>
            <strong>Sistema de produtos</strong>
            <span>Planilhas e fichas</span>
          </div>
        </div>

        <nav className="sidebar-nav" aria-label="Ferramentas">
          <button
            className={activeSection === "generator" ? "active" : ""}
            onClick={() => {
              setActiveSection("generator");
              setError("");
            }}
          >
            <Image size={19} />
            <span>
              <strong>Gerador com fotos</strong>
              <small>Selecionar produtos e criar Excel</small>
            </span>
          </button>
          <button
            className={activeSection === "filler" ? "active" : ""}
            onClick={() => {
              setActiveSection("filler");
              setError("");
            }}
          >
            <FileText size={19} />
            <span>
              <strong>Preencher planilha</strong>
              <small>Associar fichas PDF ao Excel</small>
            </span>
          </button>
          {session.user.role === "administrador" && (
            <button
              className={activeSection === "pdfAudit" ? "active" : ""}
              onClick={() => {
                setActiveSection("pdfAudit");
                setError("");
              }}
            >
              <Search size={19} />
              <span>
                <strong>Auditoria de fichas</strong>
                <small>Conferir e renomear PDFs</small>
              </span>
            </button>
          )}
          {session.user.role === "administrador" && (
            <button
              className={activeSection === "users" ? "active" : ""}
              onClick={() => {
                setActiveSection("users");
                setError("");
                loadUsers();
              }}
            >
              <Users size={19} />
              <span>
                <strong>Usuários</strong>
                <small>Contas, senhas e permissões</small>
              </span>
            </button>
          )}
        </nav>
        <div className="sidebar-session">
          <span>{session.user.username}</span>
          <small>{session.user.role}</small>
          <button className="ghost" onClick={logout}>
            <LogOut size={16} />
            Sair
          </button>
        </div>
      </aside>

      <main className="app-shell">
      <header className="topbar">
        <div>
          <h1>
            {activeSection === "generator"
              ? "Gerador de Excel com fotos"
              : activeSection === "filler"
                ? "Preenchedor de planilha"
                : activeSection === "pdfAudit"
                  ? "Auditoria de fichas PDF"
                  : "Gerenciamento de usuários"}
          </h1>
          <p>
            {activeSection === "generator"
              ? "Veja os produtos com fotos, selecione itens e gere a planilha Excel."
              : activeSection === "filler"
                ? "Envie o Excel gerado e associe as fichas PDF na ordem das abas."
                : activeSection === "pdfAudit"
                  ? "Revise as fichas sugeridas para os produtos com foto e inclua o código interno nos nomes."
                  : "Crie contas e controle senhas, status e níveis de permissão."}
          </p>
        </div>
        {activeSection === "generator" && <div className="top-actions">
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
        </div>}
      </header>

      {activeSection === "generator" && <>
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

        <button className="ghost danger" onClick={clearSelection}>
          <Trash2 size={17} />
          Limpar selecao
        </button>

        <button
          className={includePrices ? "primary" : "ghost"}
          onClick={() => {
            if (!includePrices) {
              setIncludePrices(true);
              setIncludeSheets(true);
              setPricePanelOpen(true);
            } else {
              setPricePanelOpen((current) => !current);
            }
          }}
        >
          <DollarSign size={17} />
          {includePrices
            ? pricePanelOpen ? "Ocultar preços" : "Gerenciar preços"
            : "Adicionar preço"}
        </button>

      </section>

      {includePrices && pricePanelOpen && (
        <section className="price-review-tool">
          <div className="price-review-header">
            <div>
              <strong>Preços dos produtos selecionados</strong>
              <span>{selectedPriceProducts.length} produto(s) nesta revisão</span>
            </div>
            <div className="price-review-actions">
              <button className="ghost" onClick={() => setEditedPrices({})} disabled={!selectedPriceProducts.length}>
                <RefreshCcw size={17} />
                Restaurar originais
              </button>
              <button className="ghost" onClick={() => setPricePanelOpen(false)}>
                <X size={17} />
                Fechar
              </button>
            </div>
          </div>

          <div className="price-review-summary">
            <span>Perfil: {session.user.role}</span>
            <span>Selecionados: {selectedCount}</span>
            <span>
              Preços alterados: {Object.keys(editedPrices).filter((code) => selectedCodes.includes(code)).length}
            </span>
            <span className="warn">
              {session.user.role === "vendedor"
                ? "Vendedor pode apenas manter ou aumentar"
                : "Reduções e aumentos permitidos"}
            </span>
          </div>

          <div className="price-review-list">
            <div className="price-review-row price-review-columns">
              <span>Ação</span>
              <span>Produto</span>
              <span>Preço original</span>
              <span>Novo preço</span>
              <span>Diferença</span>
            </div>
            {selectedPriceProducts.map((item) => {
              const currentPrice = Number(editedPrices[item.code] ?? item.originalPrice);
              const difference = currentPrice - Number(item.originalPrice || 0);
              return (
                <div className="price-review-row" key={item.code}>
                  <span>
                    <button
                      className="icon-button"
                      onClick={() => removeSelectedProduct(item)}
                      title="Remover da seleção"
                    >
                      <Trash2 size={15} />
                    </button>
                  </span>
                  <span className="price-product-cell">
                    <strong>{item.code} - {item.description}</strong>
                    <small>{item.brand || item.supplier || "Sem marca"}</small>
                  </span>
                  <span>{formatPrice(item.originalPrice)}</span>
                  <span>
                    {item.originalPrice == null ? (
                      <span className="warn">Preço não encontrado</span>
                    ) : (
                      <input
                        className="price-review-input"
                        type="number"
                        min={session.user.role === "vendedor" ? item.originalPrice : 0.01}
                        step="0.01"
                        value={editedPrices[item.code] ?? item.originalPrice}
                        onChange={(event) =>
                          setEditedPrices((current) => ({
                            ...current,
                            [item.code]: event.target.value,
                          }))
                        }
                      />
                    )}
                  </span>
                  <span
                    className={
                      difference < 0
                        ? "price-difference negative"
                        : difference > 0
                          ? "price-difference positive"
                          : "price-difference"
                    }
                  >
                    {Number.isFinite(difference) ? formatPrice(difference) : "-"}
                  </span>
                </div>
              );
            })}
            {!selectedPriceProducts.length && (
              <div className="price-review-empty">
                Selecione produtos no catálogo para editar os preços aqui.
              </div>
            )}
          </div>
        </section>
      )}
      </>}

      {activeSection === "filler" && (
      <section className="ordered-pdf-tool">
        <div className="ordered-pdf-controls">
          <div className="filler-control-row pdf-controls-row">
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
                setXmlPreview(null);
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

          <div className="pdf-folder-menu">
            <button
              className={showPdfFolderEditor ? "ghost active" : "ghost"}
              onClick={() => setShowPdfFolderEditor((current) => !current)}
              title="Configurar pasta padrão das fichas"
              aria-label="Configurar pasta das fichas PDF"
            >
              <MoreHorizontal size={19} />
            </button>
            {showPdfFolderEditor && (
              <div className="pdf-folder-popover">
                <label>
                  <span>Pasta das fichas PDF</span>
                  <div className="folder-input">
                    <FolderOpen size={17} />
                    <input
                      value={supplierPdfFolderPath}
                      onChange={(event) => setSupplierPdfFolderPath(event.target.value)}
                      placeholder="Pasta das fichas PDF"
                    />
                  </div>
                </label>
                <div className="pdf-folder-popover-actions">
                  <button
                    className="ghost"
                    onClick={() => setSupplierPdfFolderPath(DEFAULT_SUPPLIER_PDF_FOLDER)}
                  >
                    <RefreshCcw size={16} />
                    Restaurar padrão
                  </button>
                  <button className="primary" onClick={() => setShowPdfFolderEditor(false)}>
                    <Check size={16} />
                    Concluir
                  </button>
                </div>
              </div>
            )}
          </div>

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

          <div className="filler-control-row xml-controls-row">
            <label className="file-control">
              <FileCode2 size={17} />
              <span>{xmlFiles.length ? `${xmlFiles.length} XML(s) da NF-e` : "XMLs para buscar CEST"}</span>
              <input
                ref={xmlInputRef}
                type="file"
                accept=".xml,text/xml,application/xml"
                multiple
                onChange={(event) => {
                  setXmlFiles(Array.from(event.target.files || []));
                  setXmlPreview(null);
                  event.target.value = "";
                }}
              />
            </label>

            <button
              className="secondary"
              onClick={previewXmlCest}
              disabled={xmlLoading || !orderedWorkbookFile || !xmlFiles.length}
            >
              {xmlLoading ? <Loader2 className="spin" size={17} /> : <Search size={17} />}
              Conferir CEST
            </button>

            <button
              className="primary"
              onClick={fillWorkbookWithXmlCest}
              disabled={xmlFilling || !xmlPreview?.items?.some((item) => item.selected)}
            >
              {xmlFilling ? <Loader2 className="spin" size={17} /> : <Download size={17} />}
              Baixar Excel com CEST
            </button>
          </div>
        </div>

        {xmlPreview && (
          <div className="xml-cest-preview">
            <div className="pdf-analysis-summary">
              <span>Abas: {xmlPreview.sheetCount}</span>
              <span>XMLs: {xmlPreview.xmlCount}</span>
              <span>Itens nos XMLs: {xmlPreview.xmlItemCount}</span>
              <span className="ok">Pareados: {xmlPreview.items.filter((item) => item.selected).length}</span>
              <span className={xmlPreview.missingCestCount ? "warn" : "ok"}>
                Sem CEST: {xmlPreview.missingCestCount}
              </span>
            </div>

            <div className="xml-cest-list">
              <div className="xml-cest-row xml-cest-header">
                <span>Usar</span>
                <span>Ficha no Excel</span>
                <span>Produto no XML</span>
                <span>Correspondência</span>
                <span>EAN</span>
                <span>CEST</span>
              </div>
              {xmlPreview.items.map((item, index) => (
                <div className={`xml-cest-row ${item.cest ? "matched" : ""}`} key={`${item.sheetName}-${index}`}>
                  <span>
                    <input
                      type="checkbox"
                      checked={Boolean(item.selected)}
                      disabled={!item.cest}
                      onChange={() => toggleXmlCestItem(index)}
                    />
                  </span>
                  <span className="xml-product-cell">
                    <strong>{item.productCode || item.index}. {item.productDescription || item.sheetName}</strong>
                    <small>{item.sheetName}</small>
                  </span>
                  <span className="xml-product-cell">
                    <strong>{item.xmlDescription || "Produto não encontrado"}</strong>
                    <small>{item.xmlName ? `${item.xmlName} - item ${item.xmlItem}` : "Sem correspondência"}</small>
                  </span>
                  <span>{item.matchMethod ? `${item.matchMethod} (${item.score}%)` : "Não encontrado"}</span>
                  <span>{item.ean || item.productEan || "Sem GTIN"}</span>
                  <span className={item.cest ? "xml-cest-value" : "warn"}>{item.cest || "-"}</span>
                </div>
              ))}
            </div>
          </div>
        )}

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
      )}

      {activeSection === "pdfAudit" && session.user.role === "administrador" && (
        <section className="pdf-audit-tool">
          <div className="pdf-audit-controls">
            <label className="folder-input">
              <FolderOpen size={17} />
              <input
                value={supplierPdfFolderPath}
                onChange={(event) => setSupplierPdfFolderPath(event.target.value)}
                placeholder="Pasta das fichas PDF"
              />
            </label>
            <label className="folder-input">
              <Image size={17} />
              <input
                value={folderPath}
                onChange={(event) => setFolderPath(event.target.value)}
                placeholder="Pasta das fotos (vazio usa a padrão)"
              />
            </label>
            <button className="secondary" onClick={loadPdfAudit} disabled={pdfAuditLoading}>
              {pdfAuditLoading ? <Loader2 className="spin" size={17} /> : <Search size={17} />}
              Auditar todos
            </button>
            <button
              className="primary"
              onClick={renameAuditedPdfs}
              disabled={pdfAuditLoading || !pdfAudit?.items?.some((item) => item.selected && item.suggestedFile)}
            >
              <Pencil size={17} />
              Renomear selecionadas
            </button>
          </div>

          {pdfAuditMessage && <div className="pdf-audit-message">{pdfAuditMessage}</div>}

          {pdfAudit && (
            <>
              <div className="pdf-audit-summary">
                <span>Produtos com foto: {pdfAudit.productCount}</span>
                <span>Fichas encontradas: {pdfAudit.pdfCount}</span>
                <span className="ok">Sugestões aceitas: {pdfAudit.items.filter((item) => item.selected).length}</span>
                <span className="warn">Revisar: {pdfAudit.items.filter((item) => !item.selected).length}</span>
                <button
                  type="button"
                  className={`pdf-audit-filter ${pdfAuditOnlyCompatible ? "active" : ""}`}
                  onClick={() => setPdfAuditOnlyCompatible((value) => !value)}
                >
                  &gt; 50%
                </button>
                <label className="search-input">
                  <Search size={17} />
                  <input
                    value={pdfAuditSearch}
                    onChange={(event) => setPdfAuditSearch(event.target.value)}
                    placeholder="Buscar produto, código ou ficha"
                  />
                </label>
              </div>

              <datalist id="pdf-audit-options">
                {pdfAudit.pdfOptions.map((file) => <option value={file} key={file} />)}
              </datalist>

              <div className="pdf-audit-list">
                <div className="pdf-audit-row pdf-audit-header">
                  <span>Usar</span>
                  <span>Foto</span>
                  <span>Produto</span>
                  <span>Pontuação</span>
                  <span>Ficha sugerida ou escolhida</span>
                </div>
                {filteredPdfAuditItems.map((item) => {
                  const photoUrl = productPhotoUrl(item, folderPath);
                  return (
                    <div className={`pdf-audit-row ${item.selected ? "selected" : ""}`} key={item.productCode}>
                      <span>
                        <input
                          type="checkbox"
                          checked={Boolean(item.selected)}
                          disabled={!item.suggestedFile}
                          onChange={(event) => updatePdfAuditItem(item.auditIndex, { selected: event.target.checked })}
                        />
                      </span>
                      <span className="pdf-audit-photo">
                        {photoUrl ? <img src={photoUrl} alt="" /> : <Image size={22} />}
                      </span>
                      <span className="pdf-audit-product">
                        <strong>{item.productCode} - {item.productDescription}</strong>
                        <small>{[item.supplier, item.brand, item.factoryCode ? `Fab. ${item.factoryCode}` : ""].filter(Boolean).join(" | ")}</small>
                      </span>
                      <span className={item.score >= 62 ? "ok" : "warn"}>{item.score || 0}%</span>
                      <span>
                        <input
                          className="pdf-audit-file-input"
                          list="pdf-audit-options"
                          value={item.suggestedFile}
                          onChange={(event) => updatePdfAuditItem(item.auditIndex, {
                            suggestedFile: event.target.value,
                            selected: Boolean(event.target.value),
                          })}
                          placeholder="Escolha uma ficha PDF"
                        />
                      </span>
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {!pdfAudit && !pdfAuditLoading && (
            <div className="pdf-audit-empty">
              Informe a pasta das fichas e clique em “Auditar todos”. Nenhum arquivo será alterado nessa etapa.
            </div>
          )}
        </section>
      )}

      {activeSection === "users" && session.user.role === "administrador" && (
        <section className="users-tool">
          <form className="user-form" onSubmit={saveUser}>
            <div className="user-form-title">
              {editingUserId ? <Pencil size={19} /> : <UserPlus size={19} />}
              <strong>{editingUserId ? "Atualizar usuário" : "Novo usuário"}</strong>
            </div>
            <label>
              Usuário
              <input
                value={userForm.username}
                onChange={(event) => setUserForm((current) => ({ ...current, username: event.target.value }))}
                required
                minLength={3}
              />
            </label>
            <label>
              {editingUserId ? "Nova senha (opcional)" : "Senha"}
              <input
                type="password"
                value={userForm.password}
                onChange={(event) => setUserForm((current) => ({ ...current, password: event.target.value }))}
                required={!editingUserId}
                minLength={6}
              />
            </label>
            <label>
              Permissão
              <select
                value={userForm.role}
                onChange={(event) => setUserForm((current) => ({ ...current, role: event.target.value }))}
              >
                <option value="vendedor">Vendedor</option>
                <option value="supervisor">Supervisor</option>
                <option value="administrador">Administrador</option>
              </select>
            </label>
            {editingUserId && (
              <label className="user-active-control">
                <input
                  type="checkbox"
                  checked={userForm.active}
                  onChange={(event) => setUserForm((current) => ({ ...current, active: event.target.checked }))}
                />
                Usuário ativo
              </label>
            )}
            <div className="user-form-actions">
              <button className="primary" type="submit" disabled={usersLoading}>
                {usersLoading ? <Loader2 className="spin" size={17} /> : <UserPlus size={17} />}
                {editingUserId ? "Salvar alterações" : "Criar usuário"}
              </button>
              {editingUserId && (
                <button className="ghost" type="button" onClick={resetUserForm}>
                  <X size={17} />
                  Cancelar
                </button>
              )}
            </div>
          </form>

          <div className="users-list">
            <div className="users-list-header">
              <strong>Usuários cadastrados</strong>
              <button className="ghost" onClick={() => loadUsers()} disabled={usersLoading}>
                <RefreshCcw className={usersLoading ? "spin" : ""} size={17} />
                Atualizar
              </button>
            </div>
            <div className="table-wrap">
              <table className="users-table">
                <thead>
                  <tr>
                    <th>Usuário</th>
                    <th>Permissão</th>
                    <th>Status</th>
                    <th>Ações</th>
                  </tr>
                </thead>
                <tbody>
                  {users.map((user) => (
                    <tr key={user.id}>
                      <td className="code-cell">{user.username}</td>
                      <td className="role-cell">{user.role}</td>
                      <td>
                        <span className={user.active ? "user-status active" : "user-status inactive"}>
                          {user.active ? "ATIVO" : "INATIVO"}
                        </span>
                      </td>
                      <td className="user-actions">
                        <button className="ghost" onClick={() => editUser(user)} disabled={user.fixed}>
                          <Pencil size={16} />
                          Editar
                        </button>
                        <button
                          className="ghost danger"
                          onClick={() => removeUser(user)}
                          disabled={user.fixed || user.id === session.user.id}
                        >
                          <Trash2 size={16} />
                          Excluir
                        </button>
                      </td>
                    </tr>
                  ))}
                  {!usersLoading && users.length === 0 && (
                    <tr><td colSpan="4" className="empty-row">Nenhum usuário cadastrado.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      )}

      {error && <div className="alert">{error}</div>}

      {activeSection === "generator" && isLargeSheetGeneration && (
        <div className="alert">
          Muitos produtos selecionados para fichas individuais. Para gerar todos os produtos com velocidade,
          desmarque "Fichas individuais" e mantenha apenas a aba Produtos.
        </div>
      )}

      {activeSection === "generator" && <>
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
                <th>Preço</th>
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
                  <td>{formatPrice(item.originalPrice)}</td>
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
                  <td colSpan="7" className="empty-row">Nenhum produto carregado.</td>
                </tr>
              )}
              {loading && (
                <tr>
                  <td colSpan="7" className="empty-row">Carregando produtos...</td>
                </tr>
              )}
            </tbody>
          </table>
        </section>
      )}
      </>}
      </main>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);

import { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  X,
  Save,
  FileText,
  Package,
  Settings as SettingsIcon,
  Download,
  Upload,
  AlertCircle,
  Edit,
  Trash2,
  Copy,
  ChevronUp,
  ChevronDown,
  FileSpreadsheet,
  CheckCircle,
  XCircle,
} from 'lucide-react';
import { Product, PageContent, Settings } from '../types/catalog';
import { useLanguage } from '../hooks/useLanguage';
import { useBodyScrollLock } from '../hooks/useBodyScrollLock';
import { useAdminAuth } from '../hooks/useAdminAuth';
import { getTranslation } from '../data/translations';
import { PageTextEditor } from './PageTextEditor';
import { ProductEditorForm } from './ProductEditorForm';
import { ConfirmDialog } from './ConfirmDialog';
import { AdminLogin } from './AdminLogin';
import { importProductsFromExcel, ExcelProductRow } from '../utils/excelImport';
import { supabase } from '../lib/supabase';
import { BulkImageUpload } from './BulkImageUpload';
import { CombinedImport } from './CombinedImport';

type EditorPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  products: Product[];
  pageContent: PageContent;
  settings: Settings;
  hasUnsavedChanges: boolean;
  onUpdatePageContent: (content: PageContent) => void;
  onUpdateSettings: (settings: Settings) => void;
  onSave: () => void;
  onAddProduct: (product: Product) => Promise<void>;
  onUpdateProduct: (id: string, updates: Partial<Product>) => Promise<void>;
  onDeleteProduct: (id: string) => Promise<void>;
  onDuplicateProduct: (id: string) => Promise<void>;
  onReorderProduct: (id: string, direction: 'up' | 'down') => void;
  onResetToDefault: () => Promise<void>;
  onExportData: () => string;
  onImportData: (json: string) => Promise<{ success: boolean; error?: string }>;
  demoMode?: boolean;
};

export const EditorPanel = ({
  isOpen,
  onClose,
  products,
  pageContent,
  settings,
  hasUnsavedChanges,
  onUpdatePageContent,
  onUpdateSettings,
  onSave,
  onAddProduct,
  onUpdateProduct,
  onDeleteProduct,
  onDuplicateProduct,
  onReorderProduct,
  onResetToDefault,
  onExportData,
  onImportData,
  demoMode = false,
}: EditorPanelProps) => {
  const { language } = useLanguage();
  useBodyScrollLock(isOpen);
  const { user, isAdmin, loading, error, login, logout } = useAdminAuth();

  const [activeTab, setActiveTab] = useState<'pageText' | 'products' | 'settings' | 'importExport'>(
    'pageText'
  );
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [isAddingProduct, setIsAddingProduct] = useState(false);
  const [confirmDialog, setConfirmDialog] = useState<{
    isOpen: boolean;
    type: 'delete' | 'reset';
    productId?: string;
  }>({ isOpen: false, type: 'delete' });
  const [importJson, setImportJson] = useState('');
  const [importMessage, setImportMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(
    null
  );
  const [saveMessage, setSaveMessage] = useState(false);
  
  // Excel import state
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelProducts, setExcelProducts] = useState<(ExcelProductRow & { matchStatus: 'matched' | 'notFound' })[]>([]);
  const [excelLoading, setExcelLoading] = useState(false);
  const [excelMessage, setExcelMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const handleSave = () => {
    onSave();
    setSaveMessage(true);
    setTimeout(() => setSaveMessage(false), 3000);
  };

  const handleExport = () => {
    const data = onExportData();
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `flower-center-catalog-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImport = async () => {
    const result = await onImportData(importJson);
    if (result.success) {
      setImportMessage({ type: 'success', text: getTranslation('importSuccess', language) });
      setImportJson('');
    } else {
      setImportMessage({
        type: 'error',
        text: result.error || getTranslation('importError', language),
      });
    }
    setTimeout(() => setImportMessage(null), 5000);
  };

  const handleDeleteProduct = (id: string) => {
    setConfirmDialog({ isOpen: true, type: 'delete', productId: id });
  };

  const handleResetToDefault = () => {
    setConfirmDialog({ isOpen: true, type: 'reset' });
  };

  const confirmAction = () => {
    if (confirmDialog.type === 'delete' && confirmDialog.productId) {
      onDeleteProduct(confirmDialog.productId);
    } else if (confirmDialog.type === 'reset') {
      onResetToDefault();
    }
  };

  const handleExcelFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelFile(file);
    setExcelLoading(true);
    setExcelMessage(null);

    try {
      const excelData = await importProductsFromExcel(file);

      // Match with existing products
      const matchedData = excelData.map((excelProduct) => {
        const existingProduct = products.find(
          (p) => p.item_number.toUpperCase() === excelProduct.item_number.toUpperCase()
        );

        return {
          ...excelProduct,
          matchStatus: existingProduct ? ('matched' as const) : ('notFound' as const),
        };
      });

      setExcelProducts(matchedData);
      setExcelLoading(false);
    } catch (error) {
      setExcelMessage({
        type: 'error',
        text: error instanceof Error ? error.message : 'Failed to read Excel file',
      });
      setExcelLoading(false);
    }
  };

  const handleApplyExcelUpdates = async () => {
    setExcelLoading(true);
    setExcelMessage(null);

    try {
      if (!supabase) {
        throw new Error('Supabase not configured');
      }

      // Filter only matched products
      const matchedProducts = excelProducts.filter((p) => p.matchStatus === 'matched');

      if (matchedProducts.length === 0) {
        setExcelMessage({ type: 'error', text: 'No matched products to update' });
        setExcelLoading(false);
        return;
      }

      // Update products in Supabase
      const supabaseClient = supabase;
      const updatePromises = matchedProducts.map((excelProduct) =>
        supabaseClient
          .from('products')
          .update({
            price: excelProduct.price || null,
            excel_description: excelProduct.excel_description || null,
          })
          .eq('item_number', excelProduct.item_number)
      );

      const results = await Promise.all(updatePromises);

      // Check for errors
      const errors = results.filter((r) => r.error);
      if (errors.length > 0) {
        throw new Error(`Failed to update ${errors.length} products`);
      }

      setExcelMessage({
        type: 'success',
        text: getTranslation('excelPricingUpdated', language),
      });

      // Clear Excel data
      setExcelFile(null);
      setExcelProducts([]);

      // Refresh products (handled by parent component's save)
      // Trigger a save to reload data
      setTimeout(() => {
        window.location.reload();
      }, 1500);
    } catch (error) {
      setExcelMessage({
        type: 'error',
        text: error instanceof Error ? error.message : 'Failed to apply Excel updates',
      });
    } finally {
      setExcelLoading(false);
    }
  };

  if (!isOpen) return null;

  // Show login if not authenticated or not admin (skip check in demo mode)
  if (!demoMode && !loading && (!user || !isAdmin)) {
    // If user is logged in but not admin, show access denied
    if (user && !isAdmin) {
      return (
        <>
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[100]"
          />
          <div className="fixed inset-0 flex items-center justify-center z-[101] p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="glass-effect rounded-2xl p-8 max-w-md w-full text-center"
            >
              <button
                onClick={onClose}
                className={`absolute top-6 ${language === 'ar' ? 'left-6' : 'right-6'} text-warm-ivory/60 hover:text-warm-ivory transition-colors`}
              >
                <X className="w-6 h-6" />
              </button>

              <div className="w-16 h-16 bg-red-500/20 rounded-full flex items-center justify-center mx-auto mb-4">
                <AlertCircle className="w-8 h-8 text-red-400" />
              </div>

              <h2 className="text-2xl font-bold text-luxury-gold mb-2">
                {language === 'ar' ? 'وصول غير مصرح به' : 'Access Denied'}
              </h2>

              <p className="text-warm-ivory mb-6">
                {language === 'ar' 
                  ? 'ليس لديك صلاحيات الوصول إلى لوحة الإدارة.'
                  : 'You do not have admin access.'}
              </p>

              <button
                onClick={async () => {
                  await logout();
                  onClose();
                }}
                className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all"
              >
                {language === 'ar' ? 'حسناً' : 'OK'}
              </button>
            </motion.div>
          </div>
        </>
      );
    }

    // Show login form
    return (
      <AdminLogin
        isAdmin={isAdmin}
        loading={loading}
        error={error}
        onLogin={login}
        onLogout={logout}
        onClose={onClose}
        userEmail={user?.email}
      />
    );
  }

  return (
    <AnimatePresence>
      {isOpen && (
        <>
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[100]"
          />

          <motion.div
            initial={{ x: language === 'ar' ? '-100%' : '100%' }}
            animate={{ x: 0 }}
            exit={{ x: language === 'ar' ? '-100%' : '100%' }}
            transition={{ duration: 0.5, ease: [0.22, 1, 0.36, 1] }}
            className={`fixed top-0 ${
              language === 'ar' ? 'left-0' : 'right-0'
            } h-full w-full md:w-[600px] lg:w-[700px] glass-effect z-[101] overflow-y-auto scrollbar-hide`}
          >
            {/* Header */}
            <div className="sticky top-0 glass-effect p-6 border-b border-luxury-gold/20 z-10">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-2xl font-bold text-luxury-gold">
                  {getTranslation('editorTitle', language)}
                </h2>
                <button
                  onClick={onClose}
                  className="text-warm-ivory/60 hover:text-warm-ivory transition-colors"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>

              {/* Tabs */}
              <div className="flex gap-2 overflow-x-auto scrollbar-hide">
                {[
                  { key: 'pageText', icon: FileText },
                  { key: 'products', icon: Package },
                  { key: 'settings', icon: SettingsIcon },
                  { key: 'importExport', icon: Download },
                ].map(({ key, icon: Icon }) => (
                  <button
                    key={key}
                    onClick={() => setActiveTab(key as any)}
                    className={`flex items-center gap-2 px-4 py-2 rounded-lg font-semibold whitespace-nowrap transition-all ${
                      activeTab === key
                        ? 'bg-luxury-gold text-deep-black'
                        : 'text-warm-ivory hover:bg-warm-ivory/10'
                    }`}
                  >
                    <Icon className="w-4 h-4" />
                    {getTranslation(key, language)}
                  </button>
                ))}
              </div>

              {/* Save Button & Status */}
              <div className="flex items-center gap-3 mt-4">
                <button
                  onClick={handleSave}
                  className="flex-1 px-4 py-2 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all flex items-center justify-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  {getTranslation('saveChanges', language)}
                </button>
                {hasUnsavedChanges && !saveMessage && (
                  <div className="flex items-center gap-2 text-amber-500">
                    <AlertCircle className="w-4 h-4" />
                    <span className="text-sm">{getTranslation('unsavedChanges', language)}</span>
                  </div>
                )}
                {saveMessage && (
                  <div className="flex items-center gap-2 text-green-500">
                    <span className="text-sm">{getTranslation('changesSaved', language)}</span>
                  </div>
                )}
              </div>
            </div>

            {/* Content */}
            <div className="p-6">
              {activeTab === 'pageText' && (
                <PageTextEditor pageContent={pageContent} onUpdate={onUpdatePageContent} />
              )}

              {activeTab === 'products' && (
                <div className="space-y-4">
                  {!isAddingProduct && !editingProduct && (
                    <>
                      <button
                        onClick={() => setIsAddingProduct(true)}
                        className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all"
                      >
                        {getTranslation('addNewProduct', language)}
                      </button>

                      <div className="space-y-3">
                        {products.map((product, index) => (
                          <div
                            key={product.id}
                            className="glass-effect rounded-lg p-4 flex items-center gap-4"
                          >
                            <img
                              src={product.image_url ?? ''}
                              alt={language === 'ar' ? product.name_ar : product.name_en}
                              className="w-16 h-16 object-cover rounded-lg"
                              onError={(e) => {
                                e.currentTarget.src =
                                  'data:image/svg+xml,%3Csvg width="64" height="64" xmlns="http://www.w3.org/2000/svg"%3E%3Crect width="64" height="64" fill="%23111111"/%3E%3C/svg%3E';
                              }}
                            />
                            <div className="flex-1 min-w-0">
                              <p className="font-semibold text-warm-ivory truncate">
                                {language === 'ar' ? product.name_ar : product.name_en}
                              </p>
                              <p className="text-sm text-luxury-gold">{product.item_number}</p>
                            </div>
                            <div className="flex gap-2">
                              <button
                                onClick={() => onReorderProduct(product.id, 'up')}
                                disabled={index === 0}
                                className="p-2 text-warm-ivory hover:text-luxury-gold disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                              >
                                <ChevronUp className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => onReorderProduct(product.id, 'down')}
                                disabled={index === products.length - 1}
                                className="p-2 text-warm-ivory hover:text-luxury-gold disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                              >
                                <ChevronDown className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setEditingProduct(product)}
                                className="p-2 text-warm-ivory hover:text-luxury-gold transition-colors"
                              >
                                <Edit className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => onDuplicateProduct(product.id)}
                                className="p-2 text-warm-ivory hover:text-luxury-gold transition-colors"
                              >
                                <Copy className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => handleDeleteProduct(product.id)}
                                className="p-2 text-warm-ivory hover:text-red-500 transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </>
                  )}

                  {isAddingProduct && (
                    <ProductEditorForm
                      onSave={async (product) => {
                        await onAddProduct(product);
                        setIsAddingProduct(false);
                      }}
                      onCancel={() => setIsAddingProduct(false)}
                    />
                  )}

                  {editingProduct && (
                    <ProductEditorForm
                      product={editingProduct}
                      onSave={async (product) => {
                        await onUpdateProduct(product.id, product);
                        setEditingProduct(null);
                      }}
                      onCancel={() => setEditingProduct(null)}
                    />
                  )}
                </div>
              )}

              {activeTab === 'settings' && (
                <div className="space-y-6">
                  <div className="flex items-center justify-between">
                    <label className="text-warm-ivory font-semibold">
                      {getTranslation('showProductCodes', language)}
                    </label>
                    <input
                      type="checkbox"
                      checked={settings.showProductCodes}
                      onChange={(e) =>
                        onUpdateSettings({ ...settings, showProductCodes: e.target.checked })
                      }
                      className="w-5 h-5 bg-charcoal-black border border-luxury-gold/30 rounded accent-luxury-gold"
                    />
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-warm-ivory font-semibold">
                      {getTranslation('showDimensions', language)}
                    </label>
                    <input
                      type="checkbox"
                      checked={settings.showDimensions}
                      onChange={(e) =>
                        onUpdateSettings({ ...settings, showDimensions: e.target.checked })
                      }
                      className="w-5 h-5 bg-charcoal-black border border-luxury-gold/30 rounded accent-luxury-gold"
                    />
                  </div>

                  <div className="flex items-center justify-between">
                    <label className="text-warm-ivory font-semibold">
                      {getTranslation('featuredFirst', language)}
                    </label>
                    <input
                      type="checkbox"
                      checked={settings.featuredFirst}
                      onChange={(e) =>
                        onUpdateSettings({ ...settings, featuredFirst: e.target.checked })
                      }
                      className="w-5 h-5 bg-charcoal-black border border-luxury-gold/30 rounded accent-luxury-gold"
                    />
                  </div>

                  <div>
                    <label className="block text-sm font-semibold text-luxury-gold mb-2">
                      {getTranslation('heroFeaturedProduct', language)}
                    </label>
                    <select
                      value={settings.heroFeaturedProductId || ''}
                      onChange={(e) =>
                        onUpdateSettings({
                          ...settings,
                          heroFeaturedProductId: e.target.value || null,
                        })
                      }
                      className="w-full bg-charcoal-black border border-luxury-gold/30 rounded-lg px-4 py-3 text-warm-ivory focus:outline-none focus:ring-2 focus:ring-luxury-gold/50 transition-all"
                    >
                      <option value="">{getTranslation('selectProduct', language)}</option>
                      {products.map((product) => (
                        <option key={product.id} value={product.id}>
                          {language === 'ar' ? product.name_ar : product.name_en}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              )}

              {activeTab === 'importExport' && (
                <div className="space-y-6">
                  <div>
                    <h3 className="text-lg font-semibold text-luxury-gold mb-3">
                      {getTranslation('exportData', language)}
                    </h3>
                    <button
                      onClick={handleExport}
                      className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      {getTranslation('exportButton', language)}
                    </button>
                  </div>

                  <div className="h-px bg-luxury-gold/20" />

                  <div>
                    <h3 className="text-lg font-semibold text-luxury-gold mb-3">
                      {getTranslation('importData', language)}
                    </h3>
                    <textarea
                      value={importJson}
                      onChange={(e) => setImportJson(e.target.value)}
                      rows={10}
                      placeholder={getTranslation('pasteJson', language)}
                      className="w-full bg-charcoal-black border border-luxury-gold/30 rounded-lg px-4 py-3 text-warm-ivory focus:outline-none focus:ring-2 focus:ring-luxury-gold/50 transition-all font-mono text-sm mb-3"
                    />
                    <button
                      onClick={handleImport}
                      disabled={!importJson.trim()}
                      className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      <Upload className="w-4 h-4" />
                      {getTranslation('importButton', language)}
                    </button>
                    {importMessage && (
                      <p
                        className={`mt-3 text-sm ${
                          importMessage.type === 'success' ? 'text-green-500' : 'text-red-500'
                        }`}
                      >
                        {importMessage.text}
                      </p>
                    )}
                  </div>

                  <div className="h-px bg-luxury-gold/20" />

                  {/* Excel Import Section */}
                  <div>
                    <h3 className="text-lg font-semibold text-luxury-gold mb-3 flex items-center gap-2">
                      <FileSpreadsheet className="w-5 h-5" />
                      {getTranslation('importExcelPricing', language)}
                    </h3>
                    
                    <div className="mb-4">
                      <label className="block w-full">
                        <input
                          type="file"
                          accept=".xlsx,.xls"
                          onChange={handleExcelFileChange}
                          className="hidden"
                          id="excel-file-input"
                        />
                        <div className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all flex items-center justify-center gap-2 cursor-pointer">
                          <Upload className="w-4 h-4" />
                          {getTranslation('uploadExcel', language)}
                        </div>
                      </label>
                      <p className="text-xs text-muted-beige mt-2">
                        {getTranslation('supportedFiles', language)}
                      </p>
                    </div>

                    {excelFile && (
                      <div className="bg-charcoal-black/50 rounded-lg p-3 mb-4">
                        <p className="text-sm text-warm-ivory">
                          <span className="text-luxury-gold font-semibold">
                            {language === 'ar' ? 'الملف:' : 'File:'}
                          </span>{' '}
                          {excelFile.name}
                        </p>
                      </div>
                    )}

                    {excelLoading && (
                      <div className="text-center py-4">
                        <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-luxury-gold"></div>
                      </div>
                    )}

                    {excelMessage && (
                      <div
                        className={`mb-4 p-3 rounded-lg ${
                          excelMessage.type === 'success'
                            ? 'bg-green-500/20 border border-green-500/50 text-green-400'
                            : 'bg-red-500/20 border border-red-500/50 text-red-400'
                        }`}
                      >
                        {excelMessage.text}
                      </div>
                    )}

                    {excelProducts.length > 0 && (
                      <>
                        <div className="mb-4 overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="border-b border-luxury-gold/30">
                                <th className="text-left py-2 px-2 text-luxury-gold font-semibold">
                                  {getTranslation('sheet', language)}
                                </th>
                                <th className="text-left py-2 px-2 text-luxury-gold font-semibold">
                                  {getTranslation('itemNumber', language)}
                                </th>
                                <th className="text-left py-2 px-2 text-luxury-gold font-semibold">
                                  {getTranslation('price', language)}
                                </th>
                                <th className="text-left py-2 px-2 text-luxury-gold font-semibold">
                                  {getTranslation('sizeDetails', language)}
                                </th>
                                <th className="text-center py-2 px-2 text-luxury-gold font-semibold">
                                  {getTranslation('matchStatus', language)}
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              {excelProducts.map((product, index) => (
                                <tr
                                  key={index}
                                  className="border-b border-luxury-gold/10 hover:bg-luxury-gold/5"
                                >
                                  <td className="py-2 px-2 text-warm-ivory text-xs">
                                    {product.sheet_name}
                                  </td>
                                  <td className="py-2 px-2 text-warm-ivory font-mono text-xs">
                                    {product.item_number}
                                  </td>
                                  <td className="py-2 px-2 text-warm-ivory">
                                    {product.price || '-'}
                                  </td>
                                  <td className="py-2 px-2 text-warm-ivory text-xs">
                                    {product.excel_description || '-'}
                                  </td>
                                  <td className="py-2 px-2 text-center">
                                    {product.matchStatus === 'matched' ? (
                                      <span className="inline-flex items-center gap-1 text-green-400">
                                        <CheckCircle className="w-4 h-4" />
                                        {getTranslation('matched', language)}
                                      </span>
                                    ) : (
                                      <span className="inline-flex items-center gap-1 text-red-400">
                                        <XCircle className="w-4 h-4" />
                                        {getTranslation('notFound', language)}
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        <button
                          onClick={handleApplyExcelUpdates}
                          disabled={excelLoading || excelProducts.filter((p) => p.matchStatus === 'matched').length === 0}
                          className="w-full px-4 py-3 bg-luxury-gold text-deep-black rounded-lg font-semibold hover:bg-champagne-gold transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                          <Save className="w-4 h-4" />
                          {getTranslation('applyExcelUpdates', language)}
                        </button>

                        <p className="text-xs text-muted-beige mt-2 text-center">
                          {excelProducts.filter((p) => p.matchStatus === 'matched').length}{' '}
                          {language === 'ar' ? 'منتج مطابق' : 'matched products'} /{' '}
                          {excelProducts.length} {language === 'ar' ? 'إجمالي' : 'total'}
                        </p>
                      </>
                    )}
                  </div>

                  <div className="h-px bg-luxury-gold/20" />

                  {/* Bulk Image Upload */}
                  <BulkImageUpload
                    products={products}
                    onComplete={async () => {
                      // Refresh products after image upload
                      await new Promise((resolve) => setTimeout(resolve, 500));
                      window.location.reload();
                    }}
                  />

                  <div className="h-px bg-luxury-gold/20" />

                  {/* Combined Import */}
                  <CombinedImport
                    products={products}
                    onComplete={async () => {
                      // Refresh products after combined import
                      await new Promise((resolve) => setTimeout(resolve, 500));
                      window.location.reload();
                    }}
                  />

                  <div className="h-px bg-luxury-gold/20" />

                  <div>
                    <h3 className="text-lg font-semibold text-luxury-gold mb-3">
                      {getTranslation('resetToDefault', language)}
                    </h3>
                    <button
                      onClick={handleResetToDefault}
                      className="w-full px-4 py-3 bg-red-900/30 border border-red-500/50 text-red-400 rounded-lg font-semibold hover:bg-red-900/50 transition-all"
                    >
                      {getTranslation('resetToDefault', language)}
                    </button>
                  </div>
                </div>
              )}
            </div>
          </motion.div>

          <ConfirmDialog
            isOpen={confirmDialog.isOpen}
            onClose={() => setConfirmDialog({ ...confirmDialog, isOpen: false })}
            onConfirm={confirmAction}
            title={getTranslation('confirm', language)}
            message={getTranslation(
              confirmDialog.type === 'delete' ? 'confirmDelete' : 'confirmReset',
              language
            )}
          />
        </>
      )}
    </AnimatePresence>
  );
};

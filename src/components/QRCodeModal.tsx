import React from "react";
import { motion, AnimatePresence } from "motion/react";
import { QrCode, X, RefreshCw, CheckCircle2, Smartphone } from "lucide-react";

interface QRCodeModalProps {
  isOpen: boolean;
  onClose: () => void;
  qrUrl?: string | null;
  status?: string;
  onRefresh?: () => void;
}

export const QRCodeModal: React.FC<QRCodeModalProps> = ({
  isOpen,
  onClose,
  qrUrl,
  status,
  onRefresh,
}) => {
  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div
        id="qr-code-modal-backdrop"
        className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4 overflow-y-auto"
        onClick={(e) => {
          if (e.target === e.currentTarget) onClose();
        }}
      >
        <motion.div
          id="qr-code-modal-content"
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          className="bg-white rounded-3xl shadow-2xl max-w-md w-full overflow-hidden border border-slate-100"
        >
          {/* Header */}
          <div className="bg-gradient-to-r from-emerald-600 to-teal-700 p-6 text-white relative">
            <button
              id="qr-modal-close-btn"
              onClick={onClose}
              className="absolute top-5 right-5 text-white/80 hover:text-white p-1 rounded-full hover:bg-white/10 transition"
              title="Fechar"
            >
              <X size={20} />
            </button>
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-white/20 rounded-2xl backdrop-blur-md">
                <QrCode size={24} className="text-white" />
              </div>
              <div>
                <h3 className="font-bold text-lg leading-tight">
                  Conexão WhatsApp
                </h3>
                <p className="text-xs text-emerald-100 mt-0.5">
                  Escaneie o QR Code no seu aplicativo WhatsApp
                </p>
              </div>
            </div>
          </div>

          {/* Body */}
          <div className="p-6 flex flex-col items-center text-center">
            {qrUrl ? (
              <div className="space-y-4 w-full flex flex-col items-center">
                <div className="p-4 bg-white border-2 border-emerald-500/30 rounded-2xl shadow-inner inline-block">
                  <img
                    id="qr-code-image"
                    src={qrUrl}
                    alt="QR Code WhatsApp"
                    className="w-64 h-64 object-contain rounded-lg"
                  />
                </div>

                <div className="bg-emerald-50 text-emerald-800 text-xs font-semibold px-4 py-2.5 rounded-xl border border-emerald-200/60 flex items-center gap-2">
                  <Smartphone size={16} className="text-emerald-600 shrink-0" />
                  <span>
                    Abra o WhatsApp &gt; Aparelhos Conectados &gt; Conectar um Aparelho
                  </span>
                </div>
              </div>
            ) : (
              <div className="py-12 flex flex-col items-center justify-center space-y-4">
                <div className="w-16 h-16 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin" />
                <div className="space-y-1">
                  <p className="font-bold text-slate-800 text-sm">
                    Gerando QR Code...
                  </p>
                  <p className="text-xs text-slate-500 max-w-xs">
                    Aguarde alguns segundos enquanto o servidor inicia a sessão com o WhatsApp.
                  </p>
                </div>
              </div>
            )}

            {status && (
              <div className="mt-4 flex items-center gap-2 text-xs font-medium text-slate-600">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                <span>Status: {status}</span>
              </div>
            )}
          </div>

          {/* Footer */}
          <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-between gap-3">
            {onRefresh ? (
              <button
                id="qr-refresh-btn"
                type="button"
                onClick={onRefresh}
                className="px-3.5 py-2 text-xs font-bold text-slate-600 hover:text-slate-900 bg-white hover:bg-slate-100 border border-slate-200 rounded-xl transition flex items-center gap-1.5"
              >
                <RefreshCw size={14} />
                <span>Atualizar QR</span>
              </button>
            ) : (
              <div />
            )}
            <button
              id="qr-done-btn"
              type="button"
              onClick={onClose}
              className="px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-xl transition shadow-sm flex items-center gap-1.5"
            >
              <CheckCircle2 size={14} />
              <span>Concluído</span>
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

export default QRCodeModal;

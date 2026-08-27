import React from "react";
import { motion, AnimatePresence } from "motion/react";
import { X, QrCode, Smartphone, RefreshCw, CheckCircle2 } from "lucide-react";

interface QRCodeModalProps {
  isOpen: boolean;
  onClose: () => void;
  qrUrl?: string;
}

export function QRCodeModal({ isOpen, onClose, qrUrl }: QRCodeModalProps) {
  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div
        id="qr-code-modal-overlay"
        className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-900/70 backdrop-blur-sm p-4 overflow-y-auto"
        onClick={(e) => {
          if (e.target === e.currentTarget) {
            onClose();
          }
        }}
      >
        <motion.div
          initial={{ opacity: 0, scale: 0.92, y: 15 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.92, y: 15 }}
          transition={{ duration: 0.2, ease: "easeOut" }}
          id="qr-code-modal-card"
          className="relative bg-white w-full max-w-md rounded-3xl shadow-2xl border border-slate-100 flex flex-col overflow-hidden my-8"
        >
          {/* Header */}
          <div className="flex items-center justify-between px-6 py-5 border-b border-slate-100 bg-slate-50/50">
            <div className="flex items-center space-x-3">
              <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-2xl">
                <QrCode size={22} />
              </div>
              <div>
                <h2 className="text-base font-bold text-slate-900">
                  Escanear QR Code
                </h2>
                <p className="text-xs text-slate-500">
                  Conecte seu WhatsApp ao sistema
                </p>
              </div>
            </div>
            <button
              id="qr-code-close-btn"
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-xl transition-colors cursor-pointer"
            >
              <X size={20} />
            </button>
          </div>

          {/* Body */}
          <div className="p-6 flex flex-col items-center text-center space-y-5">
            {/* QR Code Container */}
            <div className="relative p-4 bg-white rounded-2xl border-2 border-slate-100 shadow-sm flex items-center justify-center w-64 h-64">
              {qrUrl ? (
                <img
                  src={qrUrl}
                  alt="QR Code WhatsApp"
                  className="w-full h-full object-contain rounded-lg"
                  referrerPolicy="no-referrer"
                />
              ) : (
                <div className="flex flex-col items-center justify-center text-slate-400 space-y-3">
                  <RefreshCw size={32} className="animate-spin text-emerald-500" />
                  <span className="text-xs font-medium text-slate-500">
                    Gerando QR Code...
                  </span>
                </div>
              )}
            </div>

            {/* Instructions */}
            <div className="w-full bg-slate-50 p-4 rounded-2xl border border-slate-100 text-left space-y-2.5">
              <p className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Smartphone size={14} className="text-emerald-600" />
                Como conectar:
              </p>
              <ol className="text-xs text-slate-600 space-y-1.5 list-decimal list-inside pl-1 leading-relaxed">
                <li>Abra o <strong>WhatsApp</strong> no seu celular</li>
                <li>Toque em <strong>Mais opções</strong> (Android) ou <strong>Configurações</strong> (iOS)</li>
                <li>Selecione <strong>Aparelhos conectados</strong></li>
                <li>Toque em <strong>Conectar um aparelho</strong></li>
                <li>Aponte a câmera para a tela e escaneie o código acima</li>
              </ol>
            </div>

            <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-100 px-3.5 py-2 rounded-xl w-full">
              <CheckCircle2 size={16} className="shrink-0 text-emerald-600" />
              <span>O status atualizará para <strong>Online</strong> assim que conectado.</span>
            </div>
          </div>

          {/* Footer */}
          <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end">
            <button
              id="qr-code-modal-dismiss-btn"
              onClick={onClose}
              className="px-5 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl font-semibold text-xs transition shadow-sm cursor-pointer"
            >
              Fechar
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
}

export default QRCodeModal;

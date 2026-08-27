import React from "react";
import { motion, AnimatePresence } from "motion/react";
import { X, Smartphone, AlertCircle } from "lucide-react";

interface QRCodeModalProps {
  isOpen: boolean;
  onClose: () => void;
  qrUrl: string | undefined;
}

export default function QRCodeModal({ isOpen, onClose, qrUrl }: QRCodeModalProps) {
  return (
    <AnimatePresence>
      {isOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center bg-slate-900/80 backdrop-blur-sm p-4">
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 20 }}
            className="bg-white rounded-3xl shadow-2xl overflow-hidden max-w-sm w-full relative flex flex-col items-center p-8 text-center"
          >
            {/* Close Button */}
            <button
              onClick={onClose}
              className="absolute top-4 right-4 p-2 bg-slate-50 hover:bg-slate-100 text-slate-500 rounded-full transition-colors"
            >
              <X size={20} />
            </button>

            {/* Icon */}
            <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-2xl flex items-center justify-center mb-5 rotate-3">
              <Smartphone size={32} />
            </div>

            <h3 className="text-xl font-black text-slate-800 tracking-tight mb-2">
              Conecte seu WhatsApp
            </h3>
            
            <p className="text-sm text-slate-500 mb-6 leading-relaxed">
              Abra o WhatsApp no seu celular, acesse <strong>Aparelhos Conectados</strong> e aponte a câmera para o código abaixo:
            </p>

            {/* QR Code Container */}
            <div className="p-4 bg-white border-2 border-slate-100 rounded-3xl shadow-sm mb-6">
              {qrUrl ? (
                <img
                  src={qrUrl}
                  alt="QR Code do WhatsApp"
                  className="w-56 h-56 object-contain"
                />
              ) : (
                <div className="w-56 h-56 flex flex-col items-center justify-center bg-slate-50 text-slate-400 gap-3 rounded-2xl">
                  <div className="w-8 h-8 border-4 border-slate-200 border-t-emerald-500 rounded-full animate-spin" />
                  <span className="text-xs font-bold uppercase tracking-wider">Gerando QR...</span>
                </div>
              )}
            </div>

            {/* Info Footer */}
            <div className="w-full bg-blue-50 border border-blue-100 p-4 rounded-2xl flex items-start gap-3 text-left">
              <AlertCircle className="text-blue-600 shrink-0 mt-0.5" size={16} />
              <div className="text-xs text-blue-800 leading-relaxed">
                <span className="font-bold">Atenção:</span> O QR Code expira rapidamente. Se necessário, aguarde ele recarregar automaticamente.
              </div>
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
}

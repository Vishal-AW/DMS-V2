import * as React from "react";

import { IFrameDialog } from '@pnp/spfx-controls-react/lib/IFrameDialog';

export interface IFrameDialogProps {
  url: string;
  isOpen: boolean;
  dismissPanel: (value: boolean) => void;
}


const IFrameDialogPopup: React.FunctionComponent<IFrameDialogProps> = ({ url, isOpen, dismissPanel }) => {
  const iFrameLoad = (iframe: HTMLIFrameElement) => {
    try {
      const doc = iframe.contentDocument || iframe.contentWindow?.document;
      if (!doc) return;

      // Make iframe take full available height dynamically
      iframe.style.height = "375px";
      iframe.style.width = "100%";
      iframe.style.boxShadow = "none";
      iframe.style.border = "none";

      const applyStyles = () => {
        // Inject card-style CSS
        let style = doc.getElementById('custom-share-style') as HTMLStyleElement;
        if (!style) {
          style = doc.createElement('style');
          style.id = 'custom-share-style';
          doc.head.appendChild(style);
        }

        style.innerHTML = `
                        body, #app, .od-Share-root, .ms-Panel-main, .ms-Dialog-main, div[role="dialog"] {
                          box-shadow: none !important;
                          border-radius: 8px !important;
                        }

                        /* Remove all internal shadows */
                        [style*="box-shadow"], .ms-Shadow, .ms-elevation-4, .od-Share-container {
                          box-shadow: none !important;
                          border: 1px solid #e5e5e5;
                        }

                        /* Make content area fill the card nicely */
                        .od-Share-main, main, .ms-Panel-contentInner {
                          height: 100% !important;
                          overflow-y: auto !important;
                          padding-bottom: 20px;
                        }
                      `;

        doc.querySelectorAll('button[title="Close"], [data-icon-name="Cancel"], .ms-Dialog-close, button[aria-label="Close"]')
          .forEach(el => {
            (el as HTMLElement).style.display = 'none';
          });

        doc.querySelectorAll('button[title="Manage access"], button[aria-label="Manage access"], [data-icon-name="Lock"]')
          .forEach(el => {
            (el as HTMLElement).style.display = 'none';
          });

        const header = doc.querySelector('header, .od-Share-header, [role="banner"]');
        if (header) {
          (header as HTMLElement).style.display = 'none';
        }
      };

      applyStyles();

      const observer = new MutationObserver(() => {
        console.log('Share dialog content changed - reapplying styles');
        applyStyles();
      });

      observer.observe(doc.body || doc.documentElement, {
        childList: true,
        subtree: true,
        attributes: true,
        characterData: true
      });

      // Also run again after a small delay (safety net)
      setTimeout(applyStyles, 800);
      setTimeout(applyStyles, 1500);

      // Clean up observer when iframe is unloaded (optional)
      iframe.addEventListener('unload', () => observer.disconnect());

    } catch (e) {
      console.warn('Cannot access iframe content', e);
    }
  };

  return (

    <IFrameDialog
      url={url}
      height={"auto"}
      width={"500px"}
      hidden={!isOpen}
      onDismiss={() => dismissPanel(false)}

      modalProps={{
        isBlocking: true,
        styles: {
          main: {
            maxWidth: "500px",
            width: "90vw",
            minWidth: "320px",
            maxHeight: "90vh",
            height: "auto",
            boxShadow: "none",
            borderRadius: "8px",
            border: "1px solid #e5e5e5",
            overflow: "hidden"
          }
        }
      }}

      dialogContentProps={{
        type: 2,
        showCloseButton: true,
        title: "",
        styles: {
          content: { padding: "0", overflow: "auto" },
          inner: { padding: "0" },
        }
      }}
      iframeOnLoad={iFrameLoad}
    />
  );
};

export default IFrameDialogPopup;
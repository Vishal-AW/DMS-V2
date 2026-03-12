import * as React from "react";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { useRef, useState, useEffect } from "react";
import omit from 'lodash/omit';
export interface IIFrameDialogContentProps extends React.IframeHTMLAttributes<HTMLIFrameElement> {
    close: () => void;
    iframeOnLoad?: (iframe: HTMLIFrameElement) => void;
}

/**
 * IFrame Dialog content
 */
export const IFrameDialogContent: React.FunctionComponent<IIFrameDialogContentProps> = (props: IIFrameDialogContentProps) => {
    const iframeRef = useRef<HTMLIFrameElement>(null);
    const [isContentVisible, setIsContentVisible] = useState<boolean>(false);

    const handleIFrameOnLoad = (): void => {
        try {
            const iframeElement = iframeRef.current;

            if (iframeElement && iframeElement.contentWindow) {
                const frameElement: any = iframeElement.contentWindow.frameElement;
                if (frameElement) {
                    frameElement.cancelPopUp = props.close;
                    frameElement.commitPopUp = props.close;
                    frameElement.commitPopup = props.close; // SP.UI.Dialog typo
                }
            }
        } catch (err: any) {
            if (err.name !== 'SecurityError') {
                console.error('Error accessing iframe frameElement:', err);
            }
        }

        // Call the `iframeOnLoad` prop if provided
        if (props.iframeOnLoad && iframeRef.current) {
            props.iframeOnLoad(iframeRef.current);
        }

        setIsContentVisible(true);
    };

    useEffect(() => {
        // Attach onLoad event programmatically to ensure it's properly handled
        if (iframeRef.current) {
            iframeRef.current.onload = handleIFrameOnLoad;
        }
    }, []);

    return (
        <div className="iFrameDialog">
            <iframe
                ref={iframeRef}
                {...props} // Pass additional iframe props
                frameBorder={0}
                style={{
                    width: '100%',
                    height: '100%',
                    visibility: isContentVisible ? 'visible' : 'hidden',
                }}
                {...omit(props, 'height')}
            />
            {!isContentVisible && (
                <div className="spinnerContainer">
                    <Spinner size={SpinnerSize.large} />
                </div>
            )}
        </div>
    );
};

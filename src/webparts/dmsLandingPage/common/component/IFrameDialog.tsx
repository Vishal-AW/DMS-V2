import * as React from "react";
import { useState, useEffect } from "react";
import { Dialog, IDialogContentProps, IDialogProps, IDialogStyleProps, IDialogStyles } from '@fluentui/react/lib/Dialog';
import { IFrameDialogContent } from './IFrameDialogContent';
// import * as telemetry from '../../common/telemetry';
import { Guid } from "@microsoft/sp-core-library";
import { IStyleFunctionOrObject } from "@fluentui/react/lib/Utilities";
import merge from 'lodash/merge';
import omit from 'lodash/omit';

export interface IFrameDialogProps extends IDialogProps {
  url: string;
  iframeOnLoad?: (iframe: HTMLIFrameElement) => void;
  width: string;
  height: string;
  allowFullScreen?: boolean;
  allowTransparency?: boolean;
  marginHeight?: number;
  marginWidth?: number;
  name?: string;
  sandbox?: string;
  scrolling?: string;
  seamless?: boolean;
}


const IFrameDialog: React.FunctionComponent<IFrameDialogProps> = (props: IFrameDialogProps) => {
  const [dialogId, setDialogId] = useState<string | null>(null);
  const [isStylingSet, setIsStylingSet] = useState<boolean>(false);

  useEffect(() => {
    setDialogId(`dialog-${Guid.newGuid().toString()}`);
  }, []);

  useEffect(() => {
    setDialogStyling();
  }, [props.hidden, dialogId]);

  const setDialogStyling = (): void => {
    if (!isStylingSet && !props.hidden && dialogId) {
      const element = document.querySelector(`.${dialogId} .ms-Dialog-main`) as HTMLElement;
      if (element && props.width) {
        element.style.width = props.width;
        element.style.minWidth = props.width;
        element.style.maxWidth = props.width;

        setIsStylingSet(true);
      }
    }
  };

  const {
    iframeOnLoad,
    height,
    width,
    allowFullScreen,
    allowTransparency,
    marginHeight,
    marginWidth,
    name,
    sandbox,
    scrolling,
    seamless,
    modalProps,
    className,
    dialogContentProps
  } = props;

  const dlgModalProps = {
    ...modalProps,
    onLayerDidMount: setDialogStyling,
  };

  const dlgStyles: IStyleFunctionOrObject<IDialogStyleProps, IDialogStyles> = {
    main: {
      width: width,
      maxWidth: width,
      minWidth: width,
      height: height,
    },
  };
  const dlgContentProps = merge<IDialogContentProps, IDialogContentProps, IDialogContentProps>({}, dialogContentProps as IDialogContentProps, {
    styles: {
      content: {
        display: 'flex',
        flexDirection: 'column',
        height: '100%'
      },
      inner: {
        flexGrow: 1
      },
      innerContent: {
        height: '100%'
      }
    }
  });

  return (
    <Dialog
      className={`${dialogId} ${className || ''}`}
      styles={dlgStyles}
      modalProps={dlgModalProps}
      dialogContentProps={dlgContentProps}
      {...omit(props, 'className', 'modalProps', 'dialogContentProps')}
    >
      <IFrameDialogContent
        src={props.url}
        iframeOnLoad={iframeOnLoad}
        close={props.onDismiss as () => void}
        height={height}
        allowFullScreen={allowFullScreen}
        allowTransparency={allowTransparency}
        marginHeight={marginHeight}
        marginWidth={marginWidth}
        name={name}
        sandbox={sandbox}
        scrolling={scrolling}
        seamless={seamless}
      />
    </Dialog>
  );
};

export default IFrameDialog;
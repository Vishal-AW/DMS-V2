import * as React from 'react';
import type { IDmsLandingPageProps } from "./IDmsLandingPageProps";
import { HashRouter, Route, Routes } from "react-router-dom";

import { GetAllLabel } from "../../../Services/ControlLabel";
import {
  FluentProvider,
  Toaster,
  webDarkTheme,
} from "@fluentui/react-components";
import "./styles/index.css";
import "./styles/components.css";
import "./styles/layout.css";
import "./styles/variables.css";
import "./styles/Hidedesign.css";
import { MainLayout } from './layout/main-layout';
export default class DmsLandingPage extends React.Component<IDmsLandingPageProps, {}> {
  private toasterMountRef = React.createRef<HTMLDivElement>();
  constructor(props: IDmsLandingPageProps) {
    super(props);
  }
  componentDidMount() {
    void this.getAllData();
  }

  private getAllData = async () => {
    let data: any = await GetAllLabel(this.props.context.pageContext.web.absoluteUrl, this.props.context.spHttpClient, "DefaultText");
    localStorage.setItem("DisplayLabel", JSON.stringify(data));
  };


  public render(): React.ReactElement<IDmsLandingPageProps> {
    return (
      <div style={{ width: "100vw" }}>
        <HashRouter>
          <FluentProvider theme={webDarkTheme}>
            <MainLayout context={this.props.context}>
              <div ref={this.toasterMountRef}>
                <Toaster
                  toasterId="app-toaster"
                  position="bottom-end"
                  pauseOnHover
                  mountNode={this.toasterMountRef.current}
                />
              </div>
              <Routes>
                <Route
                  path="/Dashboard"
                  element={<></>}
                />
              </Routes>
            </MainLayout>
          </FluentProvider>
        </HashRouter>
      </div>
    );
  }
}

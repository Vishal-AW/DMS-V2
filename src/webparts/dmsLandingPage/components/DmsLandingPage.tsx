import * as React from 'react';
import type { IDmsLandingPageProps } from "./IDmsLandingPageProps";
import { HashRouter, Route, Routes } from "react-router-dom";
import { GetAllLabel } from "../../../Services/ControlLabel";
import { FluentProvider, Toaster, webDarkTheme } from "@fluentui/react-components";

import "./styles/index.css";
import "./styles/components.css";
import "./styles/layout.css";
import "./styles/variables.css";
import "./styles/Hidedesign.css";
import "./styles/global.css";

import { MainLayout } from './layout/main-layout';
import Dashboard from './pages/Dashboard';
import Workspace from './pages/Workspace';
import Search from './pages/Search';
import Approvals from './pages/Approvals';
import TileSetting from './pages/TileSetting';
import TemplateMaster from './Masters/TemplateMaster';
import FolderMaster from './Masters/FolderMaster';
import ConfigMaster from './Masters/ConfigEntryForm';

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
                path="/"
                element={<MainLayout context={this.props.context}>
                  <Dashboard context={this.props.context} />
                </MainLayout>
                }
              />
              <Route path="/workspace/:workspaceId" element={<Workspace context={this.props.context} />} />
              <Route path="/search" element={<Search context={this.props.context} />} />
              <Route path="/approvals" element={<Approvals context={this.props.context} />} />
              <Route path="/tilesetting" element={<MainLayout context={this.props.context}>
                <TileSetting context={this.props.context} />
              </MainLayout>
              } />
              <Route path="/TemplateMaster"
                element={
                  <MainLayout context={this.props.context}>
                    <TemplateMaster context={this.props.context} />
                  </MainLayout>
                }
              />
              <Route path="/FolderMaster"
                element={
                  <MainLayout context={this.props.context}>
                    <FolderMaster context={this.props.context} />
                  </MainLayout>
                }
              />
              <Route path="/ConfigMaster"
                element={
                  <MainLayout context={this.props.context}>
                    <ConfigMaster context={this.props.context} />
                  </MainLayout>
                }
              />
            </Routes>
          </FluentProvider>
        </HashRouter>
      </div>
    );
  }
}

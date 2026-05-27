/* eslint-disable */
import * as React from "react";
import { useState } from "react";
import {
  Button,
} from "@fluentui/react-components";
import {
  Navigation24Regular,
} from "@fluentui/react-icons";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";


interface HeaderProps {
  onToggleSidebar: () => void;
  context: WebPartContext;
}

export function Header({ onToggleSidebar, context }: HeaderProps): JSX.Element {

  const [logoURL, setLogoURL] = useState<string>("");
  // const isMobile = window.innerWidth <= 768;

  React.useEffect(() => {
    loadLogo();
  }, []);

  const loadLogo = async (): Promise<void> => {
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const origin = new URL(webUrl).origin;
      console.log("Site URL:", webUrl);

      const apiUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/rootFolder/files?$orderby=TimeLastModified desc&$top=5`;
      console.log("Logo API URL:", apiUrl);
      const response = await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        },
      );

      if (response.ok) {
        const data = await response.json();
        const imageFile = data.value?.find(
          (file: any) =>
            file.Name.toLowerCase().endsWith(".png") ||
            file.Name.toLowerCase().endsWith(".jpg") ||
            file.Name.toLowerCase().endsWith(".jpeg") ||
            file.Name.toLowerCase().endsWith(".gif") ||
            file.Name.toLowerCase().endsWith(".svg"),
        );

        if (imageFile) {
          const logoUrl = imageFile.ServerRelativeUrl.startsWith("http")
            ? imageFile.ServerRelativeUrl
            : `${origin}${imageFile.ServerRelativeUrl}`;
          setLogoURL(logoUrl);
        }
      } else {
        console.error("Failed to load logo:", response.status, response.statusText);
      }
    } catch (error) {
      console.error("Error loading logo:", error);
    }
  };
  return (
    <header className="app-header app-header-slim" data-testid="app-header">
      <div className="header-brand">
        <Button
          appearance="subtle"
          icon={<Navigation24Regular />}
          aria-label="Toggle navigation"
          data-testid="button-sidebar-toggle"
          onClick={onToggleSidebar}
          id="toggle-btn"
        />
        <div className="header-logo"><img src={logoURL} /></div>
        {/* <span className="header-title">WorkNest</span> */}
      </div>

      <div className="header-actions">

      </div>
    </header>
  );
}

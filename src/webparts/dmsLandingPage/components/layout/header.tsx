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

// import defaultLogo from "../assets/Wizzr365Logo_20230418033130.png";

// import defaultLogo from "../../assets/defaultLogo.png";


interface HeaderProps {
  onToggleSidebar: () => void;
  context: WebPartContext;
}

export function Header({ onToggleSidebar, context }: HeaderProps): JSX.Element {

  const [logoURL, setLogoURL] = useState<string>("");
  const [allLogoData, setAllLogoData] = useState<any[]>([]);

  // const isMobile = window.innerWidth <= 768;

  React.useEffect(() => {
    loadLogo();
  }, []);

  // const loadLogo = async (): Promise<void> => {
  //   try {
  //     const webUrl = context.pageContext.web.absoluteUrl;
  //     const origin = new URL(webUrl).origin;
  //     console.log("Site URL:", webUrl);

  //     // const apiUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/rootFolder/files?$orderby=ID desc&$top=5`;
  //    // const apiUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/items?$select=ID,LogoName,Slogan,DisplaySlogan,Active,Navigation,LID,File/ServerRelativeUrl&$expand=File&$orderby=ID desc&$top=5`;

  //    const apiUrl =`${webUrl}/_api/web/lists/getByTitle('Logo')/items?$select=ID,LogoName,Slogan,DisplaySlogan,Active,Navigation,FileRef&$orderby=ID desc&$top=5`;
  //    console.log("Logo API URL:", apiUrl);
  //     const response = await context.spHttpClient.get(
  //       apiUrl,
  //       SPHttpClient.configurations.v1,
  //       {
  //         headers: {
  //           Accept: "application/json;odata=nometadata",
  //           "odata-version": "",
  //         },
  //       },
  //     );

  //     if (response.ok) {
  //       const data = await response.json();
  //       const imageFile = data.value?.find(
  //         (file: any) =>
  //           file.Name.toLowerCase().endsWith(".png") ||
  //           file.Name.toLowerCase().endsWith(".jpg") ||
  //           file.Name.toLowerCase().endsWith(".jpeg") ||
  //           file.Name.toLowerCase().endsWith(".gif") ||
  //           file.Name.toLowerCase().endsWith(".svg"),
  //       );

  //       if (imageFile) {
  //         const logoUrl = imageFile.ServerRelativeUrl.startsWith("http")
  //           ? imageFile.ServerRelativeUrl
  //           : `${origin}${imageFile.ServerRelativeUrl}`;
  //         setLogoURL(logoUrl);
  //       }
  //     } else {
  //       console.error("Failed to load logo:", response.status, response.statusText);
  //     }
  //   } catch (error) {
  //     console.error("Error loading logo:", error);
  //   }
  // };

  // const loadLogo = async (): Promise<void> => {
  //   try {
  //     const webUrl = context.pageContext.web.absoluteUrl;
  //     const origin = new URL(webUrl).origin;

  //     const apiUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/items?$select=ID,LogoName,Slogan,DisplaySlogan,Active,Navigation,FileRef,FileLeafRef&$orderby=ID desc&$top=5`;

  //     console.log("Logo API URL:", apiUrl);

  //     const response = await context.spHttpClient.get(
  //       apiUrl,
  //       SPHttpClient.configurations.v1,
  //       {
  //         headers: {
  //           Accept: "application/json;odata=nometadata",
  //           "odata-version": "",
  //         },
  //       }
  //     );

  //     if (!response.ok) {
  //       console.error("Failed to load logo:", response.status, response.statusText);
  //       return;
  //     }

  //     const data = await response.json();
  //    // setAllLogoData(data.value)
  //     const defaultLogo = "/assets/default-logo.png"; // put your default path

  //     if (data?.value && data.value.length > 0) {
  //       setAllLogoData(data.value);
  //     } else {
  //       setAllLogoData([
  //         {
  //           File: {
  //             ServerRelativeUrl: defaultLogo
  //           }
  //         }
  //       ]);
  //     }



  //     // ✅ Find image file using FileLeafRef (file name)
  //     const imageFile = data.value?.find((item: any) => {
  //       const name = item.FileLeafRef?.toLowerCase() || "";
  //       return (
  //         name.endsWith(".png") ||
  //         name.endsWith(".jpg") ||
  //         name.endsWith(".jpeg") ||
  //         name.endsWith(".gif") ||
  //         name.endsWith(".svg")
  //       );
  //     });

  //     if (imageFile) {
  //       const logoUrl = imageFile.FileRef.startsWith("http")
  //         ? imageFile.FileRef
  //         : `${origin}${imageFile.FileRef}`;

  //       console.log("Logo URL:", logoUrl);

  //       setLogoURL(logoUrl);
  //     } else {
  //       console.warn("No image file found in Logo list.");
  //     }

  //   } catch (error) {
  //     console.error("Error loading logo:", error);
  //   }
  // };


  const loadLogo = async (): Promise<void> => {
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const origin = new URL(webUrl).origin;

      const apiUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/items?$select=ID,LogoName,Slogan,DisplaySlogan,Active,Navigation,FileRef,FileLeafRef&$orderby=ID desc&$top=5`;

      const response = await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      );

      if (!response.ok) {
        console.error("Failed to load logo:", response.status, response.statusText);

        // fallback if API fails
        setAllLogoData([{ FileRef: require("../../assets/defaultLogo.png") }]);
        setLogoURL(require("../../assets/defaultLogo.png"));
        return;
      }

      const data = await response.json();

      const logoData = data?.value?.length ? data.value : [];

      setAllLogoData(logoData);

      const latestItem = logoData[0];

      if (latestItem?.FileRef) {
        const logoUrl = latestItem.FileRef.startsWith("http")
          ? latestItem.FileRef
          : `${origin}${latestItem.FileRef}`;

        setLogoURL(logoUrl);
      } else {
        // fallback
        setLogoURL(require("../../assets/defaultLogo.png"));
      }

      // Find valid image
      const imageFile = logoData.find((item: any) => {
        const name = item.FileLeafRef?.toLowerCase() || "";
        return (
          name.endsWith(".png") ||
          name.endsWith(".jpg") ||
          name.endsWith(".jpeg") ||
          name.endsWith(".gif") ||
          name.endsWith(".svg")
        );
      });

      if (imageFile && imageFile.FileRef) {
        const logoUrl = imageFile.FileRef.startsWith("http")
          ? imageFile.FileRef
          : `${origin}${imageFile.FileRef}`;

        setLogoURL(logoUrl);
      } else {
        console.warn("No image found, using default logo");

        // fallback if no image found
        setLogoURL(require("../../assets/defaultLogo.png"));
      }

    } catch (error) {
      console.error("Error loading logo:", error);

      // fallback if exception occurs
      setAllLogoData([{ FileRef: require("../../assets/defaultLogo.png") }]);
      setLogoURL(require("../../assets/defaultLogo.png"));
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
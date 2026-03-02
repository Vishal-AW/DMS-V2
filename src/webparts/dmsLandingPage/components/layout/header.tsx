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
      console.log("Loading logo from:", webUrl);

      // Method 1: Direct access to Logo document library using rootFolder/files
      try {
        const rootFolderUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/rootFolder/files?$orderby=TimeLastModified desc&$top=5`;
        console.log("Trying rootFolder/files URL:", rootFolderUrl);

        const response = await context.spHttpClient.get(
          rootFolderUrl,
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
          console.log("RootFolder files response:", data);
          if (data.value && data.value.length > 0) {
            // Find image files
            const imageFile = data.value.find(
              (file: any) =>
                file.Name.toLowerCase().endsWith(".png") ||
                file.Name.toLowerCase().endsWith(".jpg") ||
                file.Name.toLowerCase().endsWith(".jpeg") ||
                file.Name.toLowerCase().endsWith(".gif") ||
                file.Name.toLowerCase().endsWith(".svg"),
            );

            if (imageFile) {
              // Fix URL construction - ServerRelativeUrl already contains the full path
              const logoUrl = imageFile.ServerRelativeUrl.startsWith("http")
                ? imageFile.ServerRelativeUrl
                : `https://ascenworktech.sharepoint.com${imageFile.ServerRelativeUrl}`;
              console.log("Logo URL from rootFolder/files:", logoUrl);
              setLogoURL(logoUrl);
              return;
            }
          }
        } else {
          console.log(
            "RootFolder files response not ok:",
            response.status,
            response.statusText,
          );
        }
      } catch (rootFolderError) {
        console.log("RootFolder files method failed:", rootFolderError);
      }

      // Method 2: Try using getFolderByServerRelativeUrl approach
      try {
        const folderUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('/sites/SPFxApprovalManagement/Logo')/files?$orderby=TimeLastModified desc&$top=5`;
        console.log("Trying getFolderByServerRelativeUrl:", folderUrl);

        const response = await context.spHttpClient.get(
          folderUrl,
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
          console.log("Folder files response:", data);
          if (data.value && data.value.length > 0) {
            // Find image files
            const imageFile = data.value.find(
              (file: any) =>
                file.Name.toLowerCase().endsWith(".png") ||
                file.Name.toLowerCase().endsWith(".jpg") ||
                file.Name.toLowerCase().endsWith(".jpeg") ||
                file.Name.toLowerCase().endsWith(".gif") ||
                file.Name.toLowerCase().endsWith(".svg"),
            );

            if (imageFile) {
              // Fix URL construction - ServerRelativeUrl already contains the full path
              const logoUrl = imageFile.ServerRelativeUrl.startsWith("http")
                ? imageFile.ServerRelativeUrl
                : `https://ascenworktech.sharepoint.com${imageFile.ServerRelativeUrl}`;
              console.log("Logo URL from folder files:", logoUrl);
              setLogoURL(logoUrl);
              return;
            }
          }
        } else {
          console.log(
            "Folder files response not ok:",
            response.status,
            response.statusText,
          );
        }
      } catch (folderError) {
        console.log("Folder files method failed:", folderError);
      }

      // Method 3: Try to get list items with attachments from Logo list (fallback)
      try {
        const itemsUrl = `${webUrl}/_api/web/lists/getByTitle('Logo')/items?$select=Id,Title,AttachmentFiles&$expand=AttachmentFiles&$orderby=Modified desc&$top=1`;
        console.log("Trying items with attachments URL:", itemsUrl);

        const response = await context.spHttpClient.get(
          itemsUrl,
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
          console.log("Items response:", data);
          if (data.value && data.value.length > 0) {
            const item = data.value[0];
            if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
              // Find image attachments
              const imageAttachment = item.AttachmentFiles.find(
                (file: any) =>
                  file.FileName.toLowerCase().endsWith(".png") ||
                  file.FileName.toLowerCase().endsWith(".jpg") ||
                  file.FileName.toLowerCase().endsWith(".jpeg") ||
                  file.FileName.toLowerCase().endsWith(".gif") ||
                  file.FileName.toLowerCase().endsWith(".svg"),
              );

              if (imageAttachment) {
                // Fix URL construction - ServerRelativeUrl already contains the full path
                const logoUrl = imageAttachment.ServerRelativeUrl.startsWith(
                  "http",
                )
                  ? imageAttachment.ServerRelativeUrl
                  : `https://ascenworktech.sharepoint.com${imageAttachment.ServerRelativeUrl}`;
                console.log("Logo URL from attachments:", logoUrl);
                setLogoURL(logoUrl);
                return;
              }
            }
          }
        } else {
          console.log(
            "Items response not ok:",
            response.status,
            response.statusText,
          );
        }
      } catch (itemsError) {
        console.log("Items method failed:", itemsError);
      }

      // Method 4: Try common SharePoint locations as last resort
      const possibleLocations = ["Shared Documents", "SiteAssets", "Documents"];

      for (const location of possibleLocations) {
        try {
          const locUrl = `${webUrl}/_api/web/lists/getByTitle('${location}')/rootFolder/files?$filter=(substringof('logo',tolower(Name)) or substringof('approval',tolower(Name))) and (endswith(tolower(Name),'.png') or endswith(tolower(Name),'.jpg') or endswith(tolower(Name),'.jpeg') or endswith(tolower(Name),'.gif'))&$orderby=TimeLastModified desc&$top=5`;
          console.log(`Trying location: ${location}`, locUrl);

          const locResponse = await context.spHttpClient.get(
            locUrl,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "odata-version": "",
              },
            },
          );

          if (locResponse.ok) {
            const locData = await locResponse.json();
            console.log(`${location} response:`, locData);
            if (locData.value && locData.value.length > 0) {
              const logoFile = locData.value[0];
              // Fix URL construction - ServerRelativeUrl already contains the full path
              const logoUrl = logoFile.ServerRelativeUrl.startsWith("http")
                ? logoFile.ServerRelativeUrl
                : `https://ascenworktech.sharepoint.com${logoFile.ServerRelativeUrl}`;
              console.log(`Logo URL from ${location}:`, logoUrl);
              setLogoURL(logoUrl);
              return;
            }
          }
        } catch (locError) {
          console.log(`${location} method failed:`, locError);
        }
      }

      console.log("All methods failed to load logo");
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
        {/* <Menu>
                    <MenuTrigger>
                        <Button
                            appearance="subtle"
                            icon={<LocalLanguage20Regular />}
                            aria-label="Language"
                            data-testid="button-language"
                        />
                    </MenuTrigger>
                    <MenuPopover>
                        <MenuList>
                            <MenuItem
                                onClick={() => setLanguage("en")}
                                data-testid="menu-item-english"
                                style={{ fontWeight: language === "en" ? 600 : 400 }}
                            >
                                {t("language.english")}
                            </MenuItem>
                            <MenuItem
                                onClick={() => setLanguage("ar")}
                                data-testid="menu-item-arabic"
                                style={{ fontWeight: language === "ar" ? 600 : 400 }}
                            >
                                {t("language.arabic")}
                            </MenuItem>
                        </MenuList>
                    </MenuPopover>
                </Menu> */}

        {/* <Menu>
                    <MenuTrigger>
                        <Button
                            appearance="subtle"
                            icon={navMode === "sidebar" ? <PanelLeftContract20Regular /> : <TextAlignJustify20Regular />}
                            aria-label="Navigation mode"
                            data-testid="button-nav-mode"
                        />
                    </MenuTrigger>
                    <MenuPopover>
                        <MenuList>
                            <MenuItem
                                onClick={() => setNavMode("sidebar")}
                                data-testid="menu-item-sidebar"
                                style={{ fontWeight: navMode === "sidebar" ? 600 : 400 }}
                            >
                                <PanelLeftContract20Regular style={{ marginRight: "8px" }} />
                                Left Sidebar
                            </MenuItem>
                            <MenuItem
                                onClick={() => setNavMode("topnav")}
                                data-testid="menu-item-topnav"
                                style={{ fontWeight: navMode === "topnav" ? 600 : 400 }}
                            >
                                <TextAlignJustify20Regular style={{ marginRight: "8px" }} />
                                Top Navigation
                            </MenuItem>
                        </MenuList>
                    </MenuPopover>
                </Menu> */}

        {/* <Button
                    appearance="subtle"
                    icon={theme === "dark" ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
                    aria-label={theme === "dark" ? t("theme.light") : t("theme.dark")}
                    onClick={toggleTheme}
                    data-testid="button-theme-toggle"
                /> */}

        {/* <Popover
                    open={notificationsOpen}
                    onOpenChange={(_, data) => setNotificationsOpen(data.open)}
                >
                    <PopoverTrigger>
                        <div style={{ position: "relative", display: "inline-flex" }}>
                            <Button
                                appearance="subtle"
                                icon={<Alert20Regular />}
                                aria-label={t("header.notifications")}
                                data-testid="button-notifications"
                            />
                            {unreadCount > 0 && (
                                <Badge
                                    appearance="filled"
                                    color="danger"
                                    size="small"
                                    style={{
                                        position: "absolute",
                                        top: "0",
                                        right: "0",
                                        minWidth: "16px",
                                        height: "16px",
                                        fontSize: "10px",
                                        padding: "0 4px",
                                        pointerEvents: "none",
                                    }}
                                >
                                    {unreadCount}
                                </Badge>
                            )}
                        </div>
                    </PopoverTrigger>
                    <PopoverSurface
                        style={{
                            width: "320px",
                            padding: 0,
                            maxHeight: "400px",
                            overflow: "hidden",
                        }}
                    >
                        <div
                            style={{
                                padding: "var(--spacing-m)",
                                borderBottom: "1px solid var(--color-neutral-stroke-2)",
                            }}
                        >
                            <h3 style={{ margin: 0, fontSize: "var(--font-size-400)", fontWeight: 600 }}>
                                Notifications
                            </h3>
                            <p style={{ margin: 0, fontSize: "var(--font-size-200)", color: "var(--color-neutral-foreground-2)" }}>
                                {unreadCount} unread
                            </p>
                        </div>
                        <div
                            style={{
                                maxHeight: "300px",
                                overflowY: "auto",
                                padding: "var(--spacing-xs)",
                            }}
                        >
                            {notifications.length > 0 ? (
                                notifications.map((notification) => (
                                    <NotificationItem key={notification.id} notification={notification} />
                                ))
                            ) : (
                                <p
                                    style={{
                                        textAlign: "center",
                                        color: "var(--color-neutral-foreground-2)",
                                        padding: "var(--spacing-l)",
                                    }}
                                >
                                    No notifications
                                </p>
                            )}
                        </div>
                        <div
                            style={{
                                padding: "var(--spacing-xs)",
                                borderTop: "1px solid var(--color-neutral-stroke-2)",
                            }}
                        >
                            <Button
                                appearance="subtle"
                                style={{ width: "100%" }}
                                data-testid="button-view-all-notifications"
                            >
                                View all notifications
                            </Button>
                        </div>
                    </PopoverSurface>
                </Popover> */}
      </div>
    </header>
  );
}

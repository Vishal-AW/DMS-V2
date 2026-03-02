import * as React from "react";
import { ReactNode } from "react";
import { Sidebar } from "./sidebar";
import { Header } from "./header";
// import { makeStyles, tokens } from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// const useStyles = makeStyles({
//   root: {
//     minHeight: "100vh",
//     backgroundColor: tokens.colorNeutralBackground2,
//     color: tokens.colorNeutralForeground1,
//     display: "flex",
//   },
//   mainContent: {
//     flex: 1,
//     display: "flex",
//     flexDirection: "column",
//     minHeight: "100vh",
//     marginLeft: "280px", // Account for expanded sidebar
//     "@media (max-width: 768px)": {
//       marginLeft: "0px", // No margin on mobile
//     },
//     transition: "margin-left 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
//   },
//   mainContentCollapsed: {
//     marginLeft: "70px", // Collapsed sidebar width
//   },
//   main: {
//     flex: 1,
//     padding: "24px",
//     // maxWidth: "1400px",
//     // width: "100%",
//     // margin: "0 auto",
//   },
// });
interface IMainLayoutProps {
  context: WebPartContext;
  children: ReactNode;
}
function useIsMobile(breakpoint = 768): boolean {
  const [isMobile, setIsMobile] = React.useState(window.innerWidth <= breakpoint);

  React.useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth <= breakpoint);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, [breakpoint]);

  return isMobile;
}

export const MainLayout: React.FC<IMainLayoutProps> = ({
  children,
  context,
}) => {
  // const styles = useStyles();
  // const [collapsed, setCollapsed] = React.useState(false);

  // const handleToggleSidebar = () => {
  //   setCollapsed(!collapsed);
  // };

  const [sidebarCollapsed, setSidebarCollapsed] = React.useState(true);
  const [mobileMenuOpen, setMobileMenuOpen] = React.useState(false);
  // const [isMobile, setIsMobile] = React.useState(window.innerWidth <= 768);
  const isMobile = useIsMobile();


  const toggleSidebar = () => {
    if (isMobile) {
      setSidebarCollapsed(false);
      setMobileMenuOpen(!mobileMenuOpen);
    } else {
      setSidebarCollapsed(!sidebarCollapsed);
    }
  };
  React.useEffect(() => {
    if (!isMobile) {
      setMobileMenuOpen(false);
    }
  }, [isMobile]);
  const closeMobileMenu = () => setMobileMenuOpen(false);

  return (

    <div className="app-container">
      <Header onToggleSidebar={toggleSidebar} context={context} />

      <div className="app-body">
        {/* {(isMobile) && ( */}
        <>
          <div
            className={`sidebar-overlay ${mobileMenuOpen ? "visible" : ""}`}
            onClick={closeMobileMenu}
          />

          <Sidebar
            context={context}
            collapsed={sidebarCollapsed}
            mobileOpen={mobileMenuOpen}
            onCloseMobile={closeMobileMenu}
          />
        </>
        {/* )} */}

        <main className="main-content" style={{
          marginLeft: isMobile ? 0 : sidebarCollapsed ? "55px" : "290px",
          transition: "margin-left 0.3s ease"
        }}>
          <div className="content-wrapper">
            {children}
          </div>
        </main>
      </div>
    </div>
  );
};

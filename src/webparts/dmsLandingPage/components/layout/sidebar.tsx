/* eslint-disable */
import * as React from "react";
import {
  ChevronDown20Regular,
  ChevronRight20Regular,
  DocumentText24Regular,
} from "@fluentui/react-icons";
import { useState, useEffect, useRef } from "react";
import { Link, useLocation } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { loadMenuItems } from "../../../../Services/Navigation";
import * as FluentIcons from "@fluentui/react-icons";

interface NavItemProps {
  icon: JSX.Element;
  label: string;
  href?: string;
  active?: boolean;
  hasSubmenu?: boolean;
  expanded?: boolean;
  collapsed?: boolean;
  onClick?: () => void;
}

function NavItem({ icon, label, href, active, hasSubmenu, expanded, collapsed, onClick }: NavItemProps): JSX.Element {
  const content = (
    <>
      {icon}
      {!collapsed && (
        <>
          <span style={{ flex: 1, textAlign: "left" }}>{label}</span>
          {hasSubmenu && (
            expanded ? <ChevronDown20Regular /> : <ChevronRight20Regular />
          )}
        </>
      )}
    </>
  );

  if (href && !hasSubmenu) {
    return (
      <Link
        to={href}
        className={`sidebar-nav-item ${active ? "active" : ""} ${collapsed ? "collapsed" : ""}`}
        data-testid={`nav-item-${label.toLowerCase().replace(/\s+/g, "-")}`}
        title={collapsed ? label : undefined}
      >
        {content}
      </Link>
    );
  }

  return (
    <button
      className={`sidebar-nav-item ${active ? "active" : ""} ${collapsed ? "collapsed" : ""}`}
      onClick={onClick}
      data-testid={`nav-item-${label.toLowerCase().replace(/\s+/g, "-")}`}
      title={collapsed ? label : undefined}
    >
      {content}
    </button>
  );
}

interface SubNavItemProps {
  label: string;
  href: string;
  active?: boolean;
  context?: WebPartContext;
}

function SubNavItem({ context, label, href, active }: SubNavItemProps): JSX.Element {
  return (
    <Link
      to={href}
      className={`sidebar-nav-item sub-item ${active ? "active" : ""}`}
      style={{ fontSize: "var(--font-size-200)" }}
      data-testid={`nav-subitem-${label.toLowerCase().replace(/\s+/g, "-")}`}
    >
      {label}
    </Link>
  );
}



interface SidebarProps {
  context: WebPartContext;
  collapsed?: boolean;
  mobileOpen?: boolean;
  onCloseMobile?: () => void;
}



export function Sidebar({ collapsed = false, mobileOpen = false, onCloseMobile, context }: SidebarProps): JSX.Element {
  const location = useLocation();
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set());
  const [flyoutSection, setFlyoutSection] = useState<string | null>(null);
  const [dynamicNavSections, setDynamicNavSections] = useState<any[]>([]);
  const flyoutRef = useRef<HTMLDivElement>(null);

  useEffect(() => { loadMenu(); }, []);
  useEffect(() => {
    setFlyoutSection(null);
  }, [location]);

  useEffect(() => {
    if (!flyoutSection) return;

    const handleClickOutside = (event: MouseEvent) => {
      if (flyoutRef.current && !flyoutRef.current.contains(event.target as Node)) {
        const clickedNavItem = (event.target as Element).closest('.sidebar-nav-item');
        if (!clickedNavItem) {
          setFlyoutSection(null);
        }
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [flyoutSection]);

  const loadMenu = async () => {
    const data = await loadMenuItems(context);
    data && setDynamicNavSections(data);
  };

  const toggleSection = (sectionId: string) => {
    if (collapsed) {
      if (flyoutSection === sectionId) {
        setFlyoutSection(null);
      } else {
        setFlyoutSection(sectionId);
      }
      return;
    }

    if (expandedSections.has(sectionId)) {
      setExpandedSections(new Set());
    } else {
      setExpandedSections(new Set([sectionId]));
    }
  };

  const isActive = (section: { id: string; href?: string; }): boolean => {
    if (section.href === "/") {
      return location.pathname === "/";
    }
    return section.href ? location.pathname.startsWith(section.href.split("#")[0]) : false;
  };

  const isSubItemActive = (href: string): boolean => {
    const basePath = href.split("#")[0];
    return location.pathname === basePath || location.pathname.startsWith(basePath);
  };

  const handleNavClick = () => {
    setFlyoutSection(null);
    if (onCloseMobile) {
      onCloseMobile();
    }
  };



  return (
    <aside
      className={`app-sidebar ${collapsed ? "collapsed" : ""} ${mobileOpen ? "mobile-open" : ""}`}
      data-testid="app-sidebar"
    >
      <nav className="sidebar-content" style={{ marginTop: mobileOpen ? "100px" : 0 }}>
        {dynamicNavSections.map((section) => {
          const Icon = FluentIcons[section.icon as keyof typeof FluentIcons] as React.FC ?? DocumentText24Regular;

          return (
            <div key={section.id} className="sidebar-nav-section" onClick={section.href ? handleNavClick : undefined}>
              <NavItem
                icon={<Icon />}
                label={section.label}
                href={section.href}
                active={isActive(section)}
                hasSubmenu={Boolean(section.items)}
                expanded={expandedSections.has(section.id) || flyoutSection === section.id}
                collapsed={collapsed}
                onClick={section.items ? () => toggleSection(section.id) : undefined}
              />
              {!collapsed && section.items && expandedSections.has(section.id) && (
                <div className="sidebar-submenu">
                  {section.items.map((item: any) => (
                    <div key={item.label} onClick={handleNavClick}>
                      <SubNavItem
                        label={item.label}
                        href={item.href}
                        active={isSubItemActive(item.href)}
                      />
                    </div>
                  ))}
                </div>
              )}
              {collapsed && section.items && flyoutSection === section.id && (
                <div
                  className="sidebar-flyout"
                  ref={flyoutRef}
                  data-testid={`flyout-${section.id}`}
                >
                  <div className="sidebar-flyout-header">{section.label}</div>
                  {section.items.map((item: any) => (
                    <Link
                      key={item.label}
                      to={item.href}
                      className={`sidebar-flyout-item ${isSubItemActive(item.href) ? "active" : ""}`}
                      onClick={handleNavClick}
                      data-testid={`flyout-item-${item.label.toLowerCase().replace(/\s+/g, "-")}`}
                    >
                      {item.label}
                    </Link>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </nav>
    </aside>
  );
}


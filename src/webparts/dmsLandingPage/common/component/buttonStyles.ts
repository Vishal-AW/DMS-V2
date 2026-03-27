import { IButtonStyles } from "@fluentui/react";

const primaryColor = "#0078d4";
const primaryHoverColor = "#106ebe";
const neutralBorder = "#c8c6c4";
const neutralText = "#323130";

export const getPrimaryActionButtonStyles = (marginRight = 0): IButtonStyles => ({
  root: {
    minWidth: 110,
    height: 38,
    padding: "0 18px",
    marginRight,
    borderRadius: 6,
    border: `1px solid ${primaryColor}`,
    backgroundColor: primaryColor,
    fontSize: 13,
    fontWeight: 600,
  },
  rootHovered: {
    borderColor: primaryHoverColor,
    backgroundColor: primaryHoverColor,
  },
  rootPressed: {
    borderColor: primaryHoverColor,
    backgroundColor: primaryHoverColor,
  },
  rootDisabled: {
    borderColor: "#edebe9",
    backgroundColor: "#edebe9",
    color: "#a19f9d",
  },
  label: {
    fontWeight: 600,
  },
});

export const getSecondaryActionButtonStyles = (marginRight = 0): IButtonStyles => ({
  root: {
    minWidth: 110,
    height: 38,
    padding: "0 18px",
    marginRight,
    borderRadius: 6,
    border: `1px solid ${neutralBorder}`,
    backgroundColor: "#ffffff",
    color: neutralText,
    fontSize: 13,
    fontWeight: 600,
  },
  rootHovered: {
    borderColor: neutralBorder,
    backgroundColor: "#f3f2f1",
    color: neutralText,
  },
  rootPressed: {
    borderColor: neutralBorder,
    backgroundColor: "#edebe9",
    color: neutralText,
  },
  rootDisabled: {
    borderColor: "#edebe9",
    backgroundColor: "#f8f8f8",
    color: "#a19f9d",
  },
  label: {
    fontWeight: 600,
  },
});

export const getAddActionButtonStyles = (): IButtonStyles => ({
  root: {
    minWidth: 120,
    height: 38,
    padding: "0 18px",
    borderRadius: 6,
    border: `1px solid ${primaryColor}`,
    backgroundColor: "#ffffff",
    color: primaryColor,
    fontSize: 13,
    fontWeight: 600,
  },
  rootHovered: {
    borderColor: primaryColor,
    backgroundColor: "#eff6fc",
    color: primaryColor,
  },
  rootPressed: {
    borderColor: primaryColor,
    backgroundColor: "#deecf9",
    color: primaryColor,
  },
  label: {
    fontWeight: 600,
  },
});

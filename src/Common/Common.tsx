import "azure-devops-ui/Core/override.css";
import "es6-promise/auto";
import * as React from "react";
import * as ReactDOM from "react-dom";
import "./Common.scss";

export function showRootComponent(component: React.ReactElement<any>, containerId: string) {
  ReactDOM.render(component, document.getElementById(containerId));
}



export const decodeHtmlEntities = (str: string): string => {
  if (!str) return str;
  return str.replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
};

export interface Version {
  description?: string;
  questions: Question[];
};
export interface Deliverable {
  id: string;
  label: string;
  type: string;
  value: string | boolean;
};

export interface Question {
  id: string;
  text: string;
  linkedDeliverables: string[];
};

export interface Role {
  id: string;
  name: string;
  description: string;
  personName: string;
  department: string;
  email: string;
}

export interface RoleAssignment {
  roleId: string;
  raci: string;
};

export interface AnswerData {
  version: string;
  deliverables: { [deliverableId: string]: { value: string | boolean } };
  selectedQuestions: string[];
  weights: number[];
  totalWeight: number;
  uniqueResult: number;
};


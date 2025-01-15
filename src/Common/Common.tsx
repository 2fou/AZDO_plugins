import "azure-devops-ui/Core/override.css";
import "es6-promise/auto";
import * as React from "react";
import * as ReactDOM from "react-dom";
import "./Common.scss";

export function showRootComponent(component: React.ReactElement<any>, containerId: string) {
  ReactDOM.render(component, document.getElementById(containerId));
}

export interface Question {
  id: string;
  text: string;
  expectedEntries: {
    count: number;
    labels: string[];
    types: string[];
    weights: number[]; // Ensure this is defined at the entry level
  };
}

export interface EntryDetail {
  label: string;
  type: string;
  value: string | boolean;
  weight: number;
}

export interface AnswerDetail {
  questionText: string;
  entries: EntryDetail[];
  uniqueResult?: number; // Unique result per question
}

export const normalizeQuestions = (loadedQuestions: any[]): Question[] => {
  return loadedQuestions.map(question => ({
    id: question.id || "",
    text: question.text || "",
    expectedEntries: {
      count: question.expectedEntries?.count || 1,
      labels: question.expectedEntries?.labels || Array(question.expectedEntries?.count || 1).fill(""),
      types: question.expectedEntries?.types || Array(question.expectedEntries?.count || 1).fill("url"),
      weights: question.expectedEntries?.weights || Array(question.expectedEntries?.count || 1).fill(1) // Default weight
    }
  }));
};

export const decodeHtmlEntities = (str: string): string => {
  if (!str) return str;
  return str.replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
};
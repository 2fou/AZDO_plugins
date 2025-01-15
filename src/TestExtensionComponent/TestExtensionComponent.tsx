import "./testExtensionComponent.scss";
import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from "azure-devops-extension-api";
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { showRootComponent, Question, EntryDetail, normalizeQuestions, AnswerDetail, decodeHtmlEntities } from '../Common/Common';
import { WorkItemPicker } from "../Common/WorkItemPicker";

const QuestionnaireForm: React.FC = () => {
  const [questions, setQuestions] = useState<Question[]>([]);
  const [answers, setAnswers] = useState<{ [questionId: string]: AnswerDetail }>({});
  const [currentWorkItemId, setCurrentWorkItemId] = useState<string | null>(null);

  useEffect(() => {
    initializeSDK();
  }, []);

  const initializeSDK = async () => {
    console.log("Initializing SDK...");
    await SDK.init();

    try {
      const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
      const workItemId = await workItemFormService.getId();
      setCurrentWorkItemId(workItemId.toString());

      const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
      const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

      const loadedQuestions = await dataManager.getValue<Question[]>('questions', { scopeType: 'Default' }) || [];
      console.log("Loaded Questions:", loadedQuestions);
      setQuestions(normalizeQuestions(loadedQuestions));

      const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
      if (fieldValue) {
        const decodedFieldValue = decodeHtmlEntities(fieldValue);
        try {
          const parsedAnswers = JSON.parse(decodedFieldValue);
          setAnswers(parsedAnswers);
        } catch (parseError) {
          console.error("JSON parsing error:", parseError, "Field Value:", fieldValue);
        }
      }
    } catch (error) {
      console.error("SDK Initialization Error: ", error);
    }
  };

  const handleEntryChange = (
    questionId: string,
    index: number,
    label: string,
    value: string | boolean,
    type: string
  ) => {
    const question = questions.find(q => q.id === questionId);
    const weight = question ? question.expectedEntries.weights[index] : 1;

    const updatedEntries = [...(answers[questionId]?.entries || [])];
    updatedEntries[index] = { label, type, value, weight };

    setAnswers((prev) => ({
      ...prev,
      [questionId]: {
        questionText: question?.text ?? "",
        entries: updatedEntries,
      },
    }));
  };

  const handleCheckboxChange = (question: Question, checked: boolean) => {
    setAnswers((prev) => ({
      ...prev,
      [question.id]: checked
        ? createAnsweredObject(question)
        : { questionText: question.text, entries: [] },
    }));
  };

  const createAnsweredObject = (question: Question) => ({
    questionText: question.text,
    entries: question.expectedEntries.labels.map((label, i) => ({
      label,
      type: question.expectedEntries.types[i],
      value: question.expectedEntries.types[i] === 'boolean' ? false : '',
      weight: question.expectedEntries.weights[i],
    })),
  });

  const calculateUniqueResultForQuestions = () => {
    const updatedAnswers = { ...answers };

    questions.forEach((question) => {
      const answerDetail = updatedAnswers[question.id];
      if (answerDetail) {
        const entryTotal = answerDetail.entries.reduce((entrySum, entry) => {
          if (entry.value) {
            return entrySum + entry.weight;
          }
          return entrySum;
        }, 0);

        updatedAnswers[question.id].uniqueResult = entryTotal;
      }
    });

    return updatedAnswers;
  };

  const saveAnswersToWorkItemField = async () => {
    if (!currentWorkItemId) return;

    const updatedAnswers = calculateUniqueResultForQuestions();
    console.log("Updated Answers with Unique Results:", updatedAnswers);

    try {
      const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
      const serializedAnswers = JSON.stringify(updatedAnswers);
      await workItemFormService.setFieldValue('Custom.AnswersField', serializedAnswers);
      alert('Answers and unique results saved successfully to Work Item field!');
    } catch (error) {
      console.error('Error saving answers to Work Item field:', error);
    }
  };

  const renderEntryField = (entry: EntryDetail, questionId: string, index: number) => {
    switch (entry.type) {
      case 'boolean':
        return (
          <Checkbox
            label="Done"
            checked={entry.value as boolean}
            onChange={(_, checked) => handleEntryChange(questionId, index, entry.label, checked, entry.type)}
          />
        );
      case 'workItem':
        return (
          <>
            <TextField
              value={entry.value as string || ''}
              readOnly
              placeholder="Selected Work Item ID"
              className="flex-grow-text-field"
            />
            <WorkItemPicker
              onWorkItemSelected={(workItemId) => handleEntryChange(questionId, index, entry.label, workItemId.toString(), entry.type)}
            />
          </>
        );
      default:
        return (
          <TextField
            value={entry.value as string || ''}
            onChange={(_, newValue) => handleEntryChange(questionId, index, entry.label, newValue || '', entry.type)}
            placeholder="Enter URL"
            className="flex-grow-text-field"
          />
        );
    }
  };

  return (
    <div>
      {questions.length > 0 ? (
        questions.map((question) => (
          <div key={question.id}>
            <Checkbox
              label={question.text}
              checked={answers[question.id]?.entries?.length > 0}
              onChange={(_, checked) => handleCheckboxChange(question, checked)}
            />
            {answers[question.id]?.entries?.map((entry, index) => (
              <div key={`${question.id}-${entry.label}`} style={{ display: 'flex', alignItems: 'center' }}>
                <span style={{ marginRight: '8px' }}>{entry.label} :</span>
                {renderEntryField(entry, question.id, index)}
              </div>
            ))}
          </div>
        ))
      ) : (
        <div>No questions available.</div>
      )}
      <Button text="Save Answers" onClick={saveAnswersToWorkItemField} />
    </div>
  );
};

showRootComponent(<QuestionnaireForm />, 'extension-root');
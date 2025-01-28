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
  const [answers, setAnswers] = useState<{ [questionId: string]: AnswerDetail & { checked?: boolean } }>({});
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
      const normalizedQuestions = normalizeQuestions(loadedQuestions);
      setQuestions(normalizedQuestions);

      const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
      let parsedAnswers: { [questionId: string]: AnswerDetail & { checked?: boolean } } = {};
      if (fieldValue) {
        const decodedFieldValue = decodeHtmlEntities(fieldValue);
        console.log("Loaded Answer:", decodedFieldValue);
        try {
          parsedAnswers = JSON.parse(decodedFieldValue);
        } catch (parseError) {
          console.error("JSON parsing error:", parseError, "Field Value:", fieldValue);
        }
      }

      const initialAnswers = normalizedQuestions.reduce((acc, question) => {
        acc[question.id] = {
          ...(parsedAnswers[question.id] || {
            questionText: question.text,
            entries: createDefaultEntries(question),
            uniqueResult: 0,
          }),
          // Maintain the checked state if it was previously saved or set to true
          checked: parsedAnswers[question.id]?.checked || false
        };
        return acc;
      }, {} as { [questionId: string]: AnswerDetail & { checked?: boolean } });
      
      setAnswers(initialAnswers);
    } catch (error) {
      console.error("SDK Initialization Error: ", error);
    }
  };

  const createDefaultEntries = (question: Question) => 
    question.expectedEntries.labels.map((label, i) => ({
      label,
      type: question.expectedEntries.types[i],
      value: question.expectedEntries.types[i] === 'boolean' ? false : '',
      weight: question.expectedEntries.weights[i],
    }));

  const handleEntryChange = (
    questionId: string,
    index: number,
    label: string,
    value: string | boolean,
    type: string
  ) => {
    const question = questions.find(q => q.id === questionId);
    const weight = question ? question.expectedEntries.weights[index] : 1;

    setAnswers((prev) => ({
      ...prev,
      [questionId]: {
        ...prev[questionId],
        entries: prev[questionId].entries.map((entry, i) => 
          i === index ? { ...entry, value, weight } : entry
        ),
        checked: value ? true : prev[questionId].checked
      },
    }));
  };

const handleCheckboxChange = (question: Question, checked: boolean) => {
    setAnswers((prev) => ({
        ...prev,
        [question.id]: {
            ...prev[question.id],
            questionText: question.text,
            entries: checked ? (prev[question.id]?.entries.length ? prev[question.id].entries : createDefaultEntries(question)) : createDefaultEntries(question),
            uniqueResult: checked ? prev[question.id]?.uniqueResult || 0 : 0,
            checked
        },
    }));
};
  const calculateUniqueResultForQuestions = () => {
    const updatedAnswers = { ...answers };
  
    questions.forEach((question) => {
      const answerDetail = updatedAnswers[question.id] || { entries: [] };
      const entryTotal = answerDetail.entries.reduce((entrySum, entry) => {
        if (entry.value) {
          return entrySum + entry.weight;
        }
        return entrySum;
      }, 0);
  
      updatedAnswers[question.id] = {
        ...answerDetail,
        uniqueResult: entryTotal
      };
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
    const answer = answers[questionId]?.entries[index];
    switch (entry.type) {
      case 'boolean':
        return (
          <Checkbox
            label="Done"
            checked={answer?.value as boolean || false}
            onChange={(_, checked) => handleEntryChange(questionId, index, entry.label, checked, entry.type)}
          />
        );
      case 'workItem':
        return (
          <>
            <TextField
              value={answer?.value as string || ''}
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
            value={answer?.value as string || ''}
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
              checked={answers[question.id]?.checked || false}
              onChange={(_, checked) => handleCheckboxChange(question, checked)}
            />
            {answers[question.id]?.checked &&
              answers[question.id].entries.map((entry, index) => (
                <div key={`${question.id}-${entry.label}`} style={{ display: 'flex', alignItems: 'center' }}>
                  <span style={{ marginRight: '8px' }}>{entry.label} :</span>
                  {renderEntryField(entry, question.id, index)}
                </div>
              ))
            }
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
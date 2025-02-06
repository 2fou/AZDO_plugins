import "./testExtensionComponent.scss";
import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
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
  const [workItemFormService, setWorkItemFormService] = useState<IWorkItemFormService | null>(null);

  useEffect(() => {
    initializeSDK();
  }, []);

  const initializeSDK = async () => {
    console.log("Initializing SDK...");
    await SDK.init();

    try {
      const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
      setWorkItemFormService(wifService);

      const workItemId = await wifService.getId();
      setCurrentWorkItemId(workItemId.toString());

      const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
      const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

      const loadedQuestions = await dataManager.getValue<Question[]>('questions', { scopeType: 'Default' }) || [];
      console.log("Loaded Questions:", loadedQuestions);
      const normalizedQuestions = normalizeQuestions(loadedQuestions);
      setQuestions(normalizedQuestions);

      const fieldValue = await wifService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
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

  const calculateUniqueResultForQuestions = (answersToCalculate: typeof answers) => {
    const updatedAnswers = { ...answersToCalculate };
    questions.forEach((question) => {
      const answerDetail = updatedAnswers[question.id] || { entries: [] };
      const entryTotal = answerDetail.entries.reduce((sum, entry) => entry.value ? sum + entry.weight : sum, 0);
      const totalWeight = answerDetail.entries.reduce((sum, entry) => sum + entry.weight, 0);
      updatedAnswers[question.id] = { ...answerDetail, uniqueResult: entryTotal, totalWeight };
    });
    return updatedAnswers;
  };

  const handleEntryChange = (questionId: string, index: number, label: string, value: string | boolean, type: string) => {
    setAnswers((prev) => {
      const question = questions.find(q => q.id === questionId);
      const weight = question ? question.expectedEntries.weights[index] : 1;
      const newEntries = prev[questionId].entries.map((entry, i) =>
        i === index ? { ...entry, value, weight } : entry
      );
      const newAnswer = { ...prev[questionId], entries: newEntries, checked: value ? true : prev[questionId].checked };
      const newAnswers = { ...prev, [questionId]: newAnswer };
      const updatedAnswers = calculateUniqueResultForQuestions(newAnswers);

      if (workItemFormService) {
        const serialized = JSON.stringify(updatedAnswers);
        workItemFormService.setFieldValue('Custom.AnswersField', serialized);
      }
      return updatedAnswers;
    });
  };

  const handleCheckboxChange = (question: Question, checked: boolean) => {
    setAnswers((prev) => {
      const entries = checked ? (prev[question.id]?.entries.length ? prev[question.id].entries : createDefaultEntries(question)) : createDefaultEntries(question);
      const newAnswer = { ...prev[question.id], entries, uniqueResult: 0, checked };
      const newAnswers = { ...prev, [question.id]: newAnswer };
      const updatedAnswers = calculateUniqueResultForQuestions(newAnswers);

      if (workItemFormService) {
        const serialized = JSON.stringify(updatedAnswers);
        workItemFormService.setFieldValue('Custom.AnswersField', serialized);
      }
      return updatedAnswers;
    });
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
        questions.map((question) => {
          const answerDetail = answers[question.id];
          return (
            <div key={question.id}>
              <Checkbox
                label={question.text}
                checked={answerDetail?.checked || false}
                onChange={(_, checked) => handleCheckboxChange(question, checked)}
              />
              {answerDetail?.checked &&
                <>
                  {answerDetail.entries.map((entry, index) => (
                    <div key={`${question.id}-${entry.label}`} style={{ display: 'flex', alignItems: 'center' }}>
                      <span style={{ marginRight: '8px' }}>{entry.label} :</span>
                      {renderEntryField(entry, question.id, index)}
                    </div>
                  ))}
                  <div>Total Weight: {answerDetail.totalWeight}</div>
                  <div>Unique Result: {answerDetail.uniqueResult}</div>
                </>
              }
            </div>
          );
        })
      ) : (
        <div>No questions available.</div>
      )}
    </div>
  );
};

showRootComponent(<QuestionnaireForm />, 'extension-root');
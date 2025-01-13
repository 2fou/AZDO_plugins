import "./testExtensionComponent.scss";
import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from "azure-devops-extension-api";
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { showRootComponent } from '../Common/Common';
import { Question, EntryDetail, normalizeQuestions, AnswerDetail } from '../Common/Common';
import { WorkItemPicker } from "../Common/WorkItemPicker";


const QuestionnaireForm: React.FC = () => {
  const [questions, setQuestions] = useState<Question[]>([]);
  const [answers, setAnswers] = useState<{ [questionId: string]: AnswerDetail }>({});
  const [currentWorkItemId, setCurrentWorkItemId] = useState<string | null>(null);

  const decodeHtmlEntities = (str: string): string => {
    return str.replace(/&quot;/g, '"');
  };

  const initializeSDK = async () => {
    console.log("Initializing SDK...");
    await SDK.init();

    try {
      const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

      // Get the current work item ID
      const workItemId = await workItemFormService.getId();
      setCurrentWorkItemId(workItemId.toString());

      const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
      const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

      // Load questions
      const loadedQuestions = await dataManager.getValue<Question[]>('questions', { scopeType: 'Default' }) || [];
      console.log("Loaded Questions:", loadedQuestions);
      setQuestions(normalizeQuestions(loadedQuestions));

      // Load existing answers from the work item field
      const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
      console.log("Field Value before JSON.parse:", fieldValue);
      if (fieldValue) {
        // Decode the HTML entities
        const decodedFieldValue = decodeHtmlEntities(fieldValue);
        console.log("Decoded Field Value:", decodedFieldValue);
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

  useEffect(() => {
    initializeSDK();
  }, []);

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
        questionText: question?.text || "",
        entries: updatedEntries,
      },
    }));
  };

  const calculateUniqueResult = () => {
    return questions.reduce((total, question) => {
      const answerDetail = answers[question.id];
      if (answerDetail) {
        const entryTotal = answerDetail.entries.reduce((entrySum, entry) => {
          if (entry.value) {
            return entrySum + entry.weight; // Use entry-specific weight
          }
          return entrySum;
        }, 0);

        return total + entryTotal;
      }
      return total;
    }, 0);
  };

  const saveAnswersToWorkItemField = async () => {
    if (!currentWorkItemId) return;

    const uniqueResult = calculateUniqueResult();
    console.log("Unique Result:", uniqueResult);

    try {
      const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

      // Add the unique result to the serialized object
      const serializedAnswers = JSON.stringify({ ...answers, uniqueResult });

      await workItemFormService.setFieldValue('Custom.AnswersField', serializedAnswers);
      alert('Answers and unique result saved successfully to Work Item field!');
    } catch (error) {
      console.error('Error saving answers to Work Item field:', error);
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
              onChange={(e, checked) =>
                setAnswers((prev) => ({
                  ...prev,
                  [question.id]: checked
                    ? {
                      questionText: question.text,
                      entries: question.expectedEntries.labels.map((label, i) => ({
                        label,
                        type: question.expectedEntries.types[i],
                        value: question.expectedEntries.types[i] === 'boolean' ? false : '',
                        weight: question.expectedEntries.weights[i] // Set weight
                      }))
                    }
                    : { questionText: question.text, entries: [] },
                }))
              }
            />            {answers[question.id]?.entries && answers[question.id].entries.map((entry, index) => (
              <div key={index} style={{ display: 'flex', alignItems: 'center' }}>
                <span style={{ marginRight: '8px' }}>{entry.label} :</span>
                {entry.type === 'boolean' ? (
                  <Checkbox
                    label="Done"
                    checked={entry.value as boolean}
                    onChange={(e, checked) => handleEntryChange(question.id, index, entry.label, checked, entry.type)}
                  />
                ) : entry.type === 'workItem' ? (
                  <>
                    <TextField
                      value={entry.value as string || ''}
                      readOnly
                      placeholder="Selected Work Item ID"
                      className="flex-grow-text-field"
                    />
                    <WorkItemPicker
                      onWorkItemSelected={(workItemId) => handleEntryChange(question.id, index, entry.label, workItemId.toString(), entry.type)}
                    />
                  </>
                ) : (
                  <TextField
                    value={entry.value as string || ''}
                    onChange={(e, newValue) => handleEntryChange(question.id, index, entry.label, newValue || '', entry.type)}
                    placeholder="Enter URL"
                    className="flex-grow-text-field"
                  />
                )}
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

showRootComponent(<QuestionnaireForm />);
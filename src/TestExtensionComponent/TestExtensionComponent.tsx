import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from "azure-devops-extension-api";
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { showRootComponent } from '../Common/Common';

interface Question {
  id: string;
  text: string;
}

interface AnswerDetail {
  answer: boolean;
  link: string;
}

const QuestionnaireForm: React.FC = () => {
  const [questions, setQuestions] = useState<Question[]>([]);
  const [answers, setAnswers] = useState<{ [questionId: string]: AnswerDetail }>({});
  const [currentWorkItemId, setCurrentWorkItemId] = useState<string | null>(null);

  useEffect(() => {
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
        setQuestions(loadedQuestions);

        // Load existing answers from the work item field
        const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField') as string;
        if (fieldValue) {
          setAnswers(JSON.parse(fieldValue));
        }
      } catch (error) {
        console.error("SDK Initialization Error: ", error);
      }
    };

    initializeSDK();
  }, []);

  const handleAnswerChange = (questionId: string, checked: boolean) => {
    setAnswers(prev => ({
      ...prev,
      [questionId]: {
        answer: checked,
        link: checked ? (prev[questionId]?.link || '') : ''
      }
    }));
  };

  const handleLinkChange = (questionId: string, link: string) => {
    setAnswers(prev => ({
      ...prev,
      [questionId]: { ...prev[questionId], link }
    }));
  };

  const saveAnswersToWorkItemField = async () => {
    if (!currentWorkItemId) return;

    try {
      const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

      // Serialize your answers object to a JSON string or any preferred format
      const serializedAnswers = JSON.stringify(answers);

      // Assume 'Custom.AnswersField' is the reference name of your custom field
      await workItemFormService.setFieldValue('Custom.AnswersField', serializedAnswers);
      alert('Answers saved successfully to Work Item field!');
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
              checked={answers[question.id]?.answer || false}
              onChange={(e, checked) => handleAnswerChange(question.id, checked)}
            />
            {answers[question.id]?.answer && (
              <TextField
                value={answers[question.id]?.link || ''}
                onChange={(e, newValue) => handleLinkChange(question.id, newValue || '')}
                placeholder="Enter link (file server or https)"
              />
            )}
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
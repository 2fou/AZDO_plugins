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
  weight: number;
}

interface AnswerDetail {
  questionText: string;
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
        const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
        if (fieldValue) {
          setAnswers(JSON.parse(fieldValue));
        }
      } catch (error) {
        console.error("SDK Initialization Error: ", error);
      }
    };

    initializeSDK();
  }, []);

  const handleAnswerChange = (questionId: string, questionText: string, checked: boolean) => {
    setAnswers((prev) => ({
      ...prev,
      [questionId]: {
        questionText,
        answer: checked,
        link: checked ? prev[questionId]?.link || '' : '',
      },
    }));
  };
  

  const handleLinkChange = (questionId: string, link: string) => {
    setAnswers(prev => ({
      ...prev,
      [questionId]: { ...prev[questionId], link }
    }));
  };

  const calculateUniqueResult = () => {
    return questions.reduce((total, question) => {
      if (answers[question.id]?.answer) {
        return total + question.weight;
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
  
      // Ajouter le résultat unique dans l'objet sérialisé
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
              checked={answers[question.id]?.answer || false}
              onChange={(e, checked) =>  handleAnswerChange(question.id, question.text, checked)}
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
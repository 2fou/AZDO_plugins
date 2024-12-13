import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";

interface Question {
    id: string;
    text: string;
    type: string;
}

interface QuestionnaireData {
    questions: Question[];
}

interface Answers {
    [key: string]: {
        answer: boolean;
        link: string;
    };
}

const QuestionnaireForm: React.FC = () => {
    const [questionnaire, setQuestionnaire] = useState<QuestionnaireData | null>(null);
    const [answers, setAnswers] = useState<Answers>({});

    useEffect(() => {
        const initializeSDK = async () => {
            await SDK.init();
            const service = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
            const value = await service.getFieldValue('Custom.Questionnaire', { returnOriginalValue: false });
            
            if (typeof value === 'string') {
                setQuestionnaire(JSON.parse(value));
            }
        };

        initializeSDK();
    }, []);

    const handleAnswerChange = (questionId: string, checked: boolean) => {
        setAnswers(prev => ({
            ...prev,
            [questionId]: { ...prev[questionId], answer: checked }
        }));
    };

    const handleLinkChange = (questionId: string, link: string) => {
        setAnswers(prev => ({
            ...prev,
            [questionId]: { ...prev[questionId], link }
        }));
    };

    const saveAnswers = async () => {
        const service = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        await service.setFieldValue('Custom.QuestionnaireAnswers', JSON.stringify(answers));
    };

    if (!questionnaire) return <div>Loading...</div>;

    return (
        <div>
            {questionnaire.questions.map((question) => (
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
            ))}
            <Button text="Save Answers" onClick={saveAnswers} />
        </div>
    );
};

export default QuestionnaireForm;
import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { showRootComponent } from "../Common/Common";

interface Question {
    id: string;
    text: string;
    type: string;
}

interface QuestionnaireData {
    questions: Question[];
}

interface AnswerDetail {
    answer: boolean;
    link: string;
}

interface Answers {
    [key: string]: AnswerDetail;
}

const QuestionnaireForm: React.FC = () => {
    const [questionnaire, setQuestionnaire] = useState<QuestionnaireData | null>(null);
    const [answers, setAnswers] = useState<Answers>({});

    useEffect(() => {
        const initializeSDK = async () => {
            console.log("Initializing SDK...");
            await SDK.init();
            SDK.register("livrable-control", () => ({}));
            const service = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

            const questionnaireValue = await service.getFieldValue('Custom.Custom_Questionnaire', { returnOriginalValue: false });
            if (typeof questionnaireValue === 'string') {
                setQuestionnaire(JSON.parse(questionnaireValue));
                console.log("Questionnaire set:", questionnaireValue);
            } else {
                console.error("Invalid value type for questionnaire:", typeof questionnaireValue);
            }

            const answersValue = await service.getFieldValue('Custom.Custom_QuestionnaireAnswers', { returnOriginalValue: false });
            if (typeof answersValue === 'string') {
                setAnswers(JSON.parse(answersValue));
                console.log("Answers set:", answersValue);
            } else {
                console.error("Invalid value type for questionnaire answers:", typeof answersValue);
            }
        };

        initializeSDK();
    }, []);

    const handleAnswerChange = (questionId: string, checked: boolean) => {
        setAnswers(prev => ({
            ...prev,
            [questionId]: {
                answer: checked,
                link: checked ? prev[questionId]?.link : ''  // Clear link if unchecked
            }
        }));
    };
    

    const handleLinkChange = (questionId: string, link: string) => {
        setAnswers(prev => ({
            ...prev,
            [questionId]: { ...prev[questionId], link } as AnswerDetail
        }));
    };

    const saveAnswers = async () => {
        const service = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        await service.setFieldValue('Custom.Custom_QuestionnaireAnswers', JSON.stringify(answers));
        console.log("Saving answers:", answers);
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

showRootComponent(<QuestionnaireForm />);
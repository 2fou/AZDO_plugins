import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { decodeHtmlEntities,  showRootComponent } from '../Common/Common';



interface Question {
    id: string;
    text: string;
    expectedEntries: {
      count: number;
      labels: string[];
      types: string[];
      weights: number[]; // Ensure this is defined at the entry level
    };
  }
  
  interface EntryDetail {
    label: string;
    type: string;
    value: string | boolean;
    weight: number;
  }
  
  interface AnswerDetail {
    questionText: string;
    entries: EntryDetail[];
    uniqueResult?: number; // Unique result per question
    totalWeight?: number;  // Add this line
  };
  
   interface AnswerData {
    versionIndex: number;
    data: { [questionId: string]: AnswerDetail & { checked?: boolean } };
  };
  
const ProgressIndicator: React.FC = () => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [questionsProgress, setQuestionsProgress] = useState<Array<{ questionText: string, progress: number }>>([]);

    useEffect(() => {
        const initializeSDK = async () => {
            console.log("Initializing SDK...");
            await SDK.init();
            console.log("SDK Initialized!");

            try {
                setLoading(true);
                await SDK.ready();
                console.log("SDK Ready");

                const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
                console.log("WorkItemFormService obtained");

                const loadProgress = async () => {
                    const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true });
                    console.log("Field value obtained:", fieldValue);

                    if (fieldValue) {
                        try {
                            const decodedValue = decodeHtmlEntities(fieldValue as string);
                            console.log("Decoded value:", decodedValue);

                            const data = JSON.parse(decodedValue);

                            const answersData: AnswerData = data;

                            const progressData = Object.entries(answersData.data)
                                .map(([key, answer]) => mapAnswerToProgress(answer))
                                .filter(({ progress }) => progress > 0);

                            console.log("Filtered Progress data:", progressData);
                            setQuestionsProgress(progressData);
                        } catch (parseError) {
                            console.error("Error parsing answers:", parseError);
                            setError("Error parsing answers.");
                        }
                    } else {
                        setError("No field value found.");
                    }
                };

                await loadProgress();
                SDK.register(SDK.getContributionId(), async () => {
                    await loadProgress();
                });

            } catch (error) {
                console.error("Error initializing component:", error);
                setError("Failed to initialize component.");
            } finally {
                setLoading(false);
            }
        };

        initializeSDK();
    }, []);

    const mapAnswerToProgress = (answer: AnswerDetail) => {
        const answeredQuestions = answer.uniqueResult ? answer.uniqueResult.toString(2).split('1').length - 1 : 0;
        const totalBinaryWeight = answer.totalWeight ?? 0;

        const totalQuestions = totalBinaryWeight.toString(2).split('1').length - 1;
        const progress = totalQuestions > 0 ? (answeredQuestions / totalQuestions) * 100 : 0;

        // Use a label from the first entry as the question text or default to "Unknown"
        const questionText = answer.entries.length > 0 ? answer.entries[0].label : 'Unknown Question';

        return {
            questionText,
            progress
        };
    };

    const calculateColor = (progress: number) => {
        const hue = 120 * (progress / 100);  // Gradient from green to red
        return `hsl(${hue}, 100%, 50%)`;
    };

    if (loading) {
        return <div>Loading progress...</div>;
    }

    if (error) {
        return <div>Error: {error}</div>;
    }

    return (
        <div>
            {questionsProgress.map(({ questionText, progress }) => (
                <div key={questionText} style={{ marginBottom: '10px' }}>
                    <span>{questionText}: {progress.toFixed(2)}%</span>
                    <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px' }}>
                        <div
                            style={{
                                width: `${progress}%`,
                                backgroundColor: calculateColor(progress),
                                height: '24px',
                                borderRadius: '8px'
                            }}
                        />
                    </div>
                </div>
            ))}
        </div>
    );
};

showRootComponent(<ProgressIndicator />, 'progress-root');
import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { AnswerDetail, showRootComponent, decodeHtmlEntities } from '../Common/Common';

const ProgressIndicator: React.FC = () => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [questionsProgress, setQuestionsProgress] = useState<Array<{ questionText: string, progress: number, completed: number, total: number }>>([]);

    useEffect(() => {
        const initializeSDK = async () => {
            console.log("Initializing SDK...");
            await SDK.init();
            console.log("SDK Initialized!");

            try {
                setLoading(true);
                await SDK.ready();
                const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
                const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true });
                if (fieldValue) {
                    try {
                        const decodedValue = decodeHtmlEntities(fieldValue as string);
                        const answers: { [key: string]: AnswerDetail } = JSON.parse(decodedValue);

                        const progressData = Object.values(answers).map(mapAnswerToProgress);

                        setQuestionsProgress(progressData);
                    } catch (parseError) {
                        console.error("Error parsing answers:", parseError);
                        setError("Error parsing answers.");
                    }
                }
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
        if (!answer.entries) {
            console.error(`Missing entries for question ${answer.questionText}`);
            return {
                questionText: answer.questionText,
                progress: 0,
                completed: 0,
                total: 0,
            };
        }

        const totalEntries = answer.entries.length;
        const completedEntries = answer.entries.filter(e => Boolean(e.value)).length;
        const progress = totalEntries > 0 ? (completedEntries / totalEntries) * 100 : 0;

        return {
            questionText: answer.questionText,
            progress,
            completed: completedEntries,
            total: totalEntries
        };
    };

    if (loading) {
        return <div>Loading progress...</div>;
    }

    if (error) {
        return <div>Error: {error}</div>;
    }

    return (
        <div>
            <h3>Progress Indicator</h3>
            {questionsProgress.map(({ questionText, progress, completed, total }, index) => (
                <div key={questionText}>
                    <h4>{questionText}</h4>
                    <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px', marginBottom: '10px' }}>
                        <div
                            style={{
                                width: `${progress}%`,
                                backgroundColor: '#76c7c0',
                                height: '24px',
                                borderRadius: '8px'
                            }}
                        />
                    </div>
                    <div>{`Completed ${completed} out of ${total}`}</div>
                </div>
            ))}
        </div>
    );
};

showRootComponent(<ProgressIndicator />, 'progress-root');
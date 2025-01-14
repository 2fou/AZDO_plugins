import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { AnswerDetail, showRootComponent, decodeHtmlEntities } from '../Common/Common'; // Ensure the right path and structure of AnswerDetail

const calculateTotalDeliveries = (answers: { [key: string]: AnswerDetail }): number => {
    return Object.values(answers).reduce((total, entry) => total + entry.entries.length, 0);
};

const calculateCompletedDeliveries = (answers: { [key: string]: AnswerDetail }): number => {
    return Object.values(answers).reduce((total, entry) =>
        total + entry.entries.filter(e => Boolean(e.value)).length, 0
    );
};

const ProgressIndicator: React.FC = () => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [progress, setProgress] = useState<number>(0);
    const [totalDeliveries, setTotalDeliveries] = useState<number>(0);
    const [completedDeliveries, setCompletedDeliveries] = useState<number>(0);

    useEffect(() => {
        const initializeComponent = async () => {
            try {
                setLoading(true);
                await SDK.ready();
                const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
                const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: false });
                console.log("Before decoded Field Value:", fieldValue);

                if (fieldValue) {
                    try {
                        const decodedValue = decodeHtmlEntities(fieldValue as string);
                        console.log("After decoded Field Value:", decodedValue);
                        const answers: { [key: string]: AnswerDetail } = JSON.parse(decodedValue);

                        const totalDeliveries = calculateTotalDeliveries(answers);
                        const completedDeliveries = calculateCompletedDeliveries(answers);

                        setTotalDeliveries(totalDeliveries);
                        setCompletedDeliveries(completedDeliveries);

                        setProgress((completedDeliveries / totalDeliveries) * 100);
                    } catch (parseError) {
                        console.error("Error parsing answers:", parseError);
                        setError("Failed to parse answers");
                    }
                }
            } catch (error) {
                console.error("Error initializing component:", error);
                setError("Failed to initialize component");
            } finally {
                setLoading(false);
            }
        };

        initializeComponent();
    }, []);

    if (loading) {
        return <div>Loading progress...</div>;
    }

    if (error) {
        return <div>Error: {error}</div>;
    }

    return (
        <div>
            <h3>Progress Indicator</h3>
            <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px' }}>
                <div
                    style={{
                        width: `${progress}%`,
                        backgroundColor: '#76c7c0',
                        height: '24px',
                        borderRadius: '8px'
                    }}
                />
            </div>
            <div>{`Completed ${completedDeliveries} out of ${totalDeliveries} deliveries`}</div>
        </div>
    );
};

export { decodeHtmlEntities }; // Export the function

// Initialize the component
showRootComponent(<ProgressIndicator />);
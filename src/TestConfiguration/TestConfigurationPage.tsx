import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, IProjectPageService, CommonServiceIds } from "azure-devops-extension-api";


interface Question {
    id: string;
    text: string;
    weight: number;
}

const ConfigurationPage: React.FC = () => {
    const [questions, setQuestions] = useState<Question[]>([]);
    const [newQuestion, setNewQuestion] = useState('');
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        const initializeSDK = async () => {
            console.log("Initializing SDK...");
            await SDK.init();
            try {
                await loadQuestions();
            } catch (err) {
                console.error("Failed to load questions:", err);
                setError("Failed to load questions.");
            } finally {
                setIsLoading(false);
                console.log("Initialization complete.");
            }
        };

        initializeSDK();
    }, []);

    const loadQuestions = async () => {
        console.log("Loading questions...");
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        if (project) {  // Make sure project is defined
            try {
                const loadedQuestions = await dataManager.getValue<Question[]>('questions', { scopeType: 'Default' }) || [];
            if (loadedQuestions) {
                // Assigner dynamiquement les poids basés sur des puissances de 2
                const questionsWithWeights = loadedQuestions.map((question, index) => ({
                    ...question,
                    weight: Math.pow(2, index)
                }));

                console.log("Loaded Questions with Weights:", questionsWithWeights);
                setQuestions(questionsWithWeights);

            }
        } catch (error) {
            console.warn("No questions found or invalid type.");
        }
    } else {
        setError("Project not found.");
}
    };

const saveQuestions = async () => {
    try {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        if (project) {  // Make sure project is defined
            await dataManager.setValue('questions', questions, {
                scopeType: 'Default', // Use 'Default' or another correct type
            });
            alert("Questions saved successfully!");
        } else {
            setError("Project not found.");
        }
    } catch (err) {
        console.error("Failed to save questions:", err);
        setError("Failed to save questions.");
    }
};

const addQuestion = () => {
    if (newQuestion) {
        // Determine the next weight by finding the maximum weight and assigning twice that
        const nextWeight = questions.length > 0 ? Math.max(...questions.map(q => q.weight)) * 2 : 1;
        setQuestions([
            ...questions,
            { id: Date.now().toString(), text: newQuestion, weight: nextWeight }
        ]);
        setNewQuestion('');
    }
};

const removeQuestion = (id: string) => {
    setQuestions(questions.filter(q => q.id !== id));
};

if (isLoading) return <div>Loading...</div>;
if (error) return <div>{error}</div>;

return (
    <div>
        <h2>Questionnaire Configuration</h2>
        {questions.map(question => (
            <div key={question.id}>
                <TextField value={question.text} readOnly />
                <Button text="Remove" onClick={() => removeQuestion(question.id)} />
            </div>
        ))}
        <TextField
            value={newQuestion}
            onChange={(e, value) => setNewQuestion(value)}
            placeholder="Enter new question"
        />
        <Button text="Add Question" onClick={addQuestion} />
        <Button text="Save Questions" onClick={saveQuestions} />
    </div>
);
};

export default ConfigurationPage;

import React, { useState, useEffect } from 'react';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { Deliverable, Question, showRootComponent, Version } from '../Common/Common';
// Import an icon library, e.g., Material Icons
import { Icon } from 'azure-devops-ui/Icon';
import { Label } from 'office-ui-fabric-react';

const QuestionaryConfigurationPage: React.FC = () => {
    const [questions, setQuestions] = useState<Question[]>([]);
    const [newQuestionText, setNewQuestionText] = useState<string>('');
    const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
    const [versions, setVersions] = useState<Version[]>([]);
    const [currentVersionIndex, setCurrentVersionIndex] = useState<number | null>(null);
    const [currentDescription, setCurrentDescription] = useState<string>('');

    useEffect(() => {
        SDK.init();
        loadDeliverables();
        loadVersions();
    }, []);

    const getDataManager = async () => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        return extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
    };

    const loadDeliverables = async () => {
        const dataManager = await getDataManager();
        const storedDeliverables = await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || [];
        setDeliverables(storedDeliverables);
    };

    const saveVersion = async () => {
        if (questions.length === 0) {
            alert("No questions to save.");
            return;
        }

        const dataManager = await getDataManager();
        const newVersion: Version = { description: currentDescription, questions };

        if (currentVersionIndex !== null) {
            const updatedVersions = [...versions];
            updatedVersions[currentVersionIndex] = newVersion;
            setVersions(updatedVersions);
            await dataManager.setValue('questionaryVersions', updatedVersions, { scopeType: 'Default' });
        } else {
            const updatedVersions = [...versions, newVersion];
            setVersions(updatedVersions);
            setCurrentVersionIndex(updatedVersions.length - 1);
            await dataManager.setValue('questionaryVersions', updatedVersions, { scopeType: 'Default' });
        }
    };

    const loadVersions = async () => {
        const dataManager = await getDataManager();
        const storedVersions = await dataManager.getValue<Version[]>('questionaryVersions', { scopeType: 'Default' }) || [];
        setVersions(storedVersions);
        if (storedVersions.length > 0) {
            loadVersion(storedVersions.length - 1);  // Load the latest version by default
        }
    };

    const loadVersionAndRefresh = (index: number) => {
        loadVersion(index);
        refreshFromVersion(index);
    };

    const loadVersion = (index: number) => {
        if (index >= 0 && index < versions.length) {
            const version = versions[index];
            setQuestions(version.questions || []); 
            setCurrentDescription(version.description || '');
            setCurrentVersionIndex(index);
        }
    };

    const refreshFromVersion = (selectedVersionIndex: number | null) => {
        if (selectedVersionIndex !== null && selectedVersionIndex >= 0 && selectedVersionIndex < versions.length) {
            const versionQuestions = versions[selectedVersionIndex].questions;
            setQuestions(versionQuestions || []);
        }
    };

    const addQuestion = () => {
        if (newQuestionText.trim() !== '') {
            const newQuestion: Question = {
                id: Date.now().toString(),
                text: newQuestionText,
                linkedDeliverables: [],
            };
            const updatedQuestions = [...questions, newQuestion];
            setQuestions(updatedQuestions);
            setNewQuestionText('');

            if (currentVersionIndex !== null) {
                const updatedVersions = [...versions];
                updatedVersions[currentVersionIndex].questions = updatedQuestions;
                setVersions(updatedVersions);
            }
        } else {
            alert('Please enter a question text!');
        }
    };

    const updateQuestionText = (questionId: string, updatedText: string) => {
        const updatedQuestions = questions.map(q =>
            q.id === questionId ? { ...q, text: updatedText } : q
        );
        setQuestions(updatedQuestions);
    };

    const deleteQuestion = (questionId: string) => {
        setQuestions(questions.filter(q => q.id !== questionId));
    };

    const addDeliverableLink = (questionId: string, deliverableId: string) => {
        setQuestions(prevQuestions =>
            prevQuestions.map(q => q.id === questionId
                ? { ...q, linkedDeliverables: [...q.linkedDeliverables, deliverableId] }
                : q
            )
        );
    };

    const removeDeliverableLink = (questionId: string, deliverableId: string) => {
        setQuestions(prevQuestions =>
            prevQuestions.map(q => q.id === questionId
                ? { ...q, linkedDeliverables: q.linkedDeliverables.filter(id => id !== deliverableId) }
                : q
            )
        );
    };

    const deleteVersion = async (versionIndex: number) => {
        const dataManager = await getDataManager();

        if (versionIndex >= 0 && versionIndex < versions.length) {
            const updatedVersions = versions.filter((_, index) => index !== versionIndex);
            setVersions(updatedVersions);

            let newCurrentVersionIndex = null;
            if (updatedVersions.length > 0) {
                newCurrentVersionIndex = versionIndex > 0 ? versionIndex - 1 : 0;
                loadVersionAndRefresh(newCurrentVersionIndex);
            } else {
                setQuestions([]);
                setCurrentDescription('');
            }

            setCurrentVersionIndex(newCurrentVersionIndex);

            await dataManager.setValue('questionaryVersions', updatedVersions, { scopeType: 'Default' });
        } else {
            alert("Invalid version index.");
        }
    };

    return (
        <div style={{ maxWidth: '800px', margin: '0 auto', padding: '20px' }}>
            <h2>Questionary Configuration</h2>
            <div className="question-form">
                <TextField
                    value={newQuestionText}
                    onChange={(e, value) => setNewQuestionText(value || '')}
                    placeholder="Enter a new question"
                />
                <Button text="Add Question" onClick={addQuestion} />
            </div>
            <div className="version-description">
                <TextField
                    value={currentDescription}
                    onChange={(e, value) => setCurrentDescription(value || '')}
                    placeholder="Enter version description"
                />
                <Button text="Save Current Version" onClick={saveVersion} />
            </div>
            <div className="save-controls">
                <Button
                    text="Refresh Current Version"
                    onClick={() => {
                        if (currentVersionIndex !== null) {
                            refreshFromVersion(currentVersionIndex);
                        } else {
                            alert("Please select a version to refresh.");
                        }
                    }}
                />
            </div>
            <div className="delete-version">
                <Button
                    text="Delete Current Version"
                    onClick={() => {
                        if (currentVersionIndex !== null) {
                            deleteVersion(currentVersionIndex);
                        } else {
                            alert("No version selected to delete.");
                        }
                    }}
                />
            </div>
            <div className="version-selector">
                <h3>Select Version to Load</h3>
                <select
                    onChange={(e) => {
                        const selectedIndex = parseInt(e.target.value);
                        loadVersionAndRefresh(selectedIndex);
                    }}
                    defaultValue={currentVersionIndex !== null ? currentVersionIndex.toString() : 'new'}
                >
                    {versions.map((version, index) => (
                        <option key={index} value={index}>
                            {`Version ${index + 1}: ${version.description || 'No description'}`}
                        </option>
                    ))}
                    <option value="new">New Version</option>
                </select>
            </div>
            <div className="question-list">
                {questions.map(question => (
                    <div key={question.id} style={{ marginBottom: '10px', border: '1px solid #ccc', padding: '10px' }}>
                        <TextField
                            value={question.text}
                            onChange={(e, value) => updateQuestionText(question.id, value || '')}
                        />
                        <Button text="Delete" onClick={() => deleteQuestion(question.id)} />

                        <div>
                            <Label>Related Deliverables:</Label>
                            {/* Select for adding deliverables */}
                            <select data-question-id={question.id}>
                                <option value="">Select Deliverable</option>
                                {deliverables
                                    .filter(d => !question.linkedDeliverables.includes(d.id))
                                    .map(deliverable => (
                                        <option key={deliverable.id} value={deliverable.id}>
                                            {deliverable.label}
                                        </option>
                                    ))}
                            </select>
                            <Button
                                text="Add"
                                onClick={() => {
                                    const select = document.querySelector(`select[data-question-id="${question.id}"]`);
                                    const value = (select as HTMLSelectElement)?.value;
                                    if (value) {
                                        addDeliverableLink(question.id, value);
                                    }
                                }}
                            />

                            {/* List and remove added deliverables */}
                            <div>
                                {question.linkedDeliverables.map(deliverableId => (
                                    <div key={deliverableId} style={{ display: 'flex', alignItems: 'center' }}>
                                        <span>
                                            {deliverables.find(d => d.id === deliverableId)?.label}
                                        </span>
                                        {/* Replace remove button with an icon */}
                                        <Icon
                                            iconName='Delete'
                                            ariaLabel="Remove"
                                            onClick={() => removeDeliverableLink(question.id, deliverableId)}
                                            style={{ marginLeft: '8px', cursor: 'pointer', color: '#ff0000' }}
                                        />
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                ))}
            </div>
        </div>
    );
};

showRootComponent(<QuestionaryConfigurationPage />, 'questionary-configuration-root');
export default QuestionaryConfigurationPage;
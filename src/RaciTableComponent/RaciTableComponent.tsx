import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { showRootComponent, decodeHtmlEntities, fetchRoles } from '../Common/Common';
import './RaciTableComponent.scss';

interface Role {
    id: string;
    name: string;
}

interface RoleAssignmentData {
    [questionId: string]: {
        [entryLabel: string]: { roleId: string; raci: string }[];
    };
}

const RaciTableComponent: React.FC = () => {
    const [roleAssignments, setRoleAssignments] = useState<RoleAssignmentData>({});
    const [roles, setRoles] = useState<Role[]>([]);

    useEffect(() => {
        const initializeComponent = async () => {
            await SDK.init();
            const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

            try {
                const rolesData = await fetchRoles();
                setRoles(rolesData);

                const roleAssignmentsField = await wifService.getFieldValue('Custom.RoleAssignmentsField', { returnOriginalValue: true }) as string;
                if (roleAssignmentsField) {
                    const decodedRoleAssignments = decodeHtmlEntities(roleAssignmentsField);
                    const parsedAssignments = JSON.parse(decodedRoleAssignments);
                    setRoleAssignments(parsedAssignments);
                }
            } catch (error) {
                console.error('Error loading RACI assignments:', error);
            }
        };

        initializeComponent();
    }, []);

    const renderRaciLabels = (assignments: { roleId: string; raci: string }[], roleId: string) => {
        const assignment = assignments.find(a => a.roleId === roleId);
        if (assignment) {
            return (
                <>
                    {assignment.raci.split('').map(label => (
                        <span key={label} className={`raci-${label}`}>{label}</span>
                    ))}
                </>
            );
        }
        return null;
    };

    return (
        <table className="raci-table">
            <thead>
                <tr>
                    <th>Entry</th>
                    {roles.map(role => (
                        <th key={role.id}>{role.name}</th>
                    ))}
                </tr>
            </thead>
            <tbody>
                {Object.entries(roleAssignments).map(([questionId, entries]) => 
                    Object.entries(entries).map(([entryLabel, assignments]) => (
                        <tr key={`${questionId}-${entryLabel}`}>
                            <td>{entryLabel}</td>
                            {roles.map(role => (
                                <td key={role.id}>
                                    {renderRaciLabels(assignments, role.id)}
                                </td>
                            ))}
                        </tr>
                    ))
                )}
            </tbody>
        </table>
    );
};

showRootComponent(<RaciTableComponent />, 'raci-table-root');

export default RaciTableComponent;
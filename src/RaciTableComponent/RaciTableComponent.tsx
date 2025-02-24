import React, { useEffect, useState } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';
import { Deliverable, Role, decodeHtmlEntities, showRootComponent } from '../Common/Common';
import './RaciTableComponent.scss'; // Assuming you have a separate CSS file for styles

const RaciTableComponent: React.FC = () => {
    const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
    const [roles, setRoles] = useState<Role[]>([]);
    const [roleAssignments, setRoleAssignments] = useState<{ [deliverableId: string]: { roleId: string, raci: string }[] }>({});

    useEffect(() => {
        const initializeComponent = async () => {
            try {
                await SDK.init();
                const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

                const rolesData = await dataManager.getValue<Role[]>('roles', { scopeType: 'Default' }) || [];
                setRoles(rolesData);

                const storedDeliverables = await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || [];
                setDeliverables(storedDeliverables);

                // Fetch and decode the role assignments
                const roleAssignmentsField = await wifService.getFieldValue('Custom.RoleAssignmentsField', { returnOriginalValue: true }) as string;
                if (roleAssignmentsField) {
                    const decodedAssignments = decodeHtmlEntities(roleAssignmentsField);
                    setRoleAssignments(JSON.parse(decodedAssignments));
                }
            } catch (error) {
                console.error("Initialization error:", error);
            }
        };
        initializeComponent();
    }, []);

    const getCellStyle = (raci: string) => {
        const baseStyle = { padding: '5px', borderRight: '1px solid #ddd', borderBottom: '1px solid #ddd' };
        if (raci.includes('R')) return { ...baseStyle, backgroundColor: '#ffcccc' }; // Red for Responsible
        if (raci.includes('A')) return { ...baseStyle, backgroundColor: '#ccffcc' }; // Green for Accountable
        if (raci.includes('C')) return { ...baseStyle, backgroundColor: '#ccccff' }; // Blue for Consulted
        if (raci.includes('I')) return { ...baseStyle, backgroundColor: '#ffffcc' }; // Yellow for Informed
        return baseStyle;
    };

    const renderMatrix = () => {
        const deliverableIds = Object.keys(roleAssignments);
        const relevantDeliverables = deliverables.filter(deliverable => deliverableIds.includes(deliverable.id));
        const roleIdsSet = new Set<string>();

        Object.values(roleAssignments).forEach(assignments => {
            assignments.forEach(assignment => roleIdsSet.add(assignment.roleId));
        });

        const relevantRoles = roles.filter(role => roleIdsSet.has(role.id));

        return (
            <table className="raci-table">
                <thead>
                    <tr>
                        <th>Deliverable</th>
                        {relevantRoles.map(role => <th key={role.id}>{role.name}</th>)}
                    </tr>
                </thead>
                <tbody>
                    {relevantDeliverables.map(deliverable => (
                        <tr key={deliverable.id}>
                            <td>{deliverable.label}</td>
                            {relevantRoles.map(role => {
                                const roleAssignment = roleAssignments[deliverable.id]?.find(ra => ra.roleId === role.id);
                                return (
                                    <td key={role.id} style={getCellStyle(roleAssignment?.raci || '')}>
                                        {roleAssignment?.raci || '-'}
                                    </td>
                                );
                            })}
                        </tr>
                    ))}
                </tbody>
            </table>
        );
    };

    return (
        <div>
            <h2>RACI Matrix</h2>
            {renderMatrix()}
        </div>
    );
};

showRootComponent(<RaciTableComponent />, 'raci-table-root');
export default RaciTableComponent;
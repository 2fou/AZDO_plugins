import * as React from 'react';
import { useState, useEffect } from 'react';
import { Button } from 'azure-devops-ui/Button';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { Role, showRootComponent } from '../Common/Common';


const RoleConfigurationPage: React.FC = () => {
    const [roles, setRoles] = useState<Role[]>([]);
    const [newRole, setNewRole] = useState<Role>({
        id: '',
        name: '',
        description: '',
        personName: '',
        department: '',
        email: ''
    });

    useEffect(() => {
        initializePage();
    }, []);

    const initializePage = async () => {
        await SDK.init();
        await loadRoles();
    };

    const loadRoles = async () => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        const storedRoles = await dataManager.getValue<Role[]>('roles', { scopeType: 'Default' }) || [];
        setRoles(storedRoles);
    };

    const saveRoles = async (updatedRoles: Role[]) => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        await dataManager.setValue('roles', updatedRoles, { scopeType: 'Default' });
    };

    const addRole = () => {
        if (newRole.name) {
            const updatedRoles = [...roles, { ...newRole, id: Date.now().toString() }];
            setRoles(updatedRoles);
            setNewRole({ id: '', name: '', description: '', personName: '', department: '', email: '' });
            saveRoles(updatedRoles);
        } else {
            alert("Role name is required!");
        }
    };

    const updateRole = (roleId: string, updatedRole: Partial<Role>) => {
        const updatedRoles = roles.map(role => role.id === roleId ? { ...role, ...updatedRole } : role);
        setRoles(updatedRoles);
        saveRoles(updatedRoles);
    };

    const deleteRole = (roleId: string) => {
        const updatedRoles = roles.filter(role => role.id !== roleId);
        setRoles(updatedRoles);
        saveRoles(updatedRoles);
    };

    return (
        <div className="role-configuration-container" style={{ maxWidth: '800px', margin: '0 auto', padding: '20px' }}>
            <h2>Role Configuration</h2>
            <div className="role-form">
                <TextField
                    value={newRole.name}
                    onChange={(e, value) => setNewRole({ ...newRole, name: value || '' })}
                    placeholder="Role Name (required)"
                    width={TextFieldWidth.standard}
                />
                <TextField
                    value={newRole.description}
                    onChange={(e, value) => setNewRole({ ...newRole, description: value || '' })}
                    placeholder="Role Description"
                    width={TextFieldWidth.standard}
                />
                <TextField
                    value={newRole.personName}
                    onChange={(e, value) => setNewRole({ ...newRole, personName: value || '' })}
                    placeholder="Person Name"
                    width={TextFieldWidth.standard}
                />
                <TextField
                    value={newRole.department}
                    onChange={(e, value) => setNewRole({ ...newRole, department: value || '' })}
                    placeholder="Department"
                    width={TextFieldWidth.standard}
                />
                <TextField
                    value={newRole.email}
                    onChange={(e, value) => setNewRole({ ...newRole, email: value || '' })}
                    placeholder="Email Address"
                    width={TextFieldWidth.standard}
                />
                <Button text="Add Role" onClick={addRole} />
            </div>
            <ul className="role-list">
                {roles.map(role => (
                    <li key={role.id}>
                        <TextField
                            value={role.name}
                            onChange={(e, value) => updateRole(role.id, { name: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <TextField
                            value={role.description}
                            onChange={(e, value) => updateRole(role.id, { description: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <TextField
                            value={role.personName}
                            onChange={(e, value) => updateRole(role.id, { personName: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <TextField
                            value={role.department}
                            onChange={(e, value) => updateRole(role.id, { department: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <TextField
                            value={role.email}
                            onChange={(e, value) => updateRole(role.id, { email: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <Button text="Delete" onClick={() => deleteRole(role.id)} />
                    </li>
                ))}
            </ul>
        </div>
    );
};

showRootComponent(<RoleConfigurationPage />, 'Role-Configuration-root');

export default RoleConfigurationPage;
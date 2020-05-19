from settings import *
import os
import requests
from collections import namedtuple
import xlsxwriter
from datetime import datetime, date


class Workspaces(object):

    def __init__(self):

        self.uri = WATCHDOX_API_BASE_URL
        r = self.__send_request__(endpoint='/sessions/create', method='POST', params=None,
                                  headers={'Content-Type': 'application/json'}, postdata=None,
                                  json={"email": WATCHDOX_API_EMAIL, "password": WATCHDOX_API_PASSWORD})
        self.ssid = r.json()['ssid']
        self.auth_header = {"Authorization": "Bearer {token}".format(token=self.ssid)}
        self.headers = {"Content-Type": "application/json", "Authorization": "Bearer {token}".format(token=self.ssid)}
        self.room_id = None
        self.base_folder_id = None
        self.base_folder_name = ""
        self.folder_id = None
        self.folder_path = ""
        self.processed_folder_path = ""
        self.rejected_folder_path = ""

    def __send_request__(self, endpoint, params=None, method=None, postdata=None, json=None, headers=None, files=None):

        if method == 'GET':
            return requests.get('{uri}{endpoint}'.format(uri=self.uri, endpoint=endpoint), headers=headers,
                                params=params).json()
        else:
            if json:
                return requests.post('{uri}{endpoint}'.format(uri=self.uri, endpoint=endpoint), headers=headers,
                                     json=json)
            elif files:
                return requests.post('{uri}{endpoint}'.format(uri=self.uri, endpoint=endpoint), headers=headers,
                                     files=files)
            else:
                return requests.post('{uri}{endpoint}'.format(uri=self.uri, endpoint=endpoint), headers=headers)

    def get_room_info(self, room_id=None):

        return self.__send_request__(endpoint='/rooms/{room_id}/info'.format(room_id=room_id), method='POST',
                                     headers=self.headers).json()

    def edit_room_name(self, room_id=None, new_room_name=None):

        room_info = {'name': new_room_name}
        return self.__send_request__(endpoint='/rooms/{room_id}/edit'.format(room_id=room_id), method='POST',
                                     headers=self.headers, json=room_info).json()

    def rename_document(self, room_id=None, document_id=None, new_document_name=None):

        document_info = {"newFileName": new_document_name}
        return self.__send_request__(
            endpoint='/rooms/{room_id}/documents/{document_id}/rename'.format(room_id=room_id, document_id=document_id),
            method='POST', headers=self.headers, json=document_info)

    def move_document(self, room_id=None, document_id=None, folder_path=None):

        document_info = {"path": folder_path}
        return self.__send_request__(
            endpoint='/rooms/{room_id}/documents/{document_id}/move'.format(room_id=room_id, document_id=document_id),
            method='POST', headers=self.headers, json=document_info)

    def move_folder(self, room_id=None, new_folder_path=None, current_folder_path=None):

        folder_info = {"newPath": new_folder_path, "currentPath": current_folder_path}
        return self.__send_request__(endpoint='/rooms/{room_id}/folders/move'.format(room_id=room_id), method='POST',
                                     headers=self.headers, json=folder_info)

    def delete_folders(self, room_id=None, folder_paths=None, folder_ids=None, folder_guids=None):

        if folder_ids:
            folder_info = {"folderIds": folder_ids, "forceAction": True, "message": None, "deviceType": "BROWSER"}
        elif folder_guids:
            folder_info = {"folderGuids": folder_guids, "forceAction": True, "message": None, "deviceType": "BROWSER"}
        else:
            folder_info = {"folderPaths": folder_paths, "forceAction": True, "message": None, "deviceType": "BROWSER"}
        print(folder_ids, folder_guids, folder_paths, folder_info)
        return self.__send_request__(endpoint='/rooms/{room_id}/documents/delete'.format(room_id=room_id),
                                     method='POST', headers=self.headers, json=folder_info).json()

    def create_folder(self, folder_id=None, room_id=None, new_folder_name=None):

        new_folder_info = {"roomId": room_id, "parentId": folder_id, "name": new_folder_name}
        return self.__send_request__(endpoint='/rooms/folders/create', method='POST', headers=self.headers,
                                     json=new_folder_info)

    def get_room_folders(self, room_id=None):

        return self.__send_request__(endpoint='/rooms/{room_id}/folders'.format(room_id=room_id), method='GET',
                                     headers=self.headers)

    def get_folder_info(self, folder_id=None, folder_path=None, room_id=None):

        if folder_id:
            folder_info = {"roomId": room_id, "folderId": folder_id}
        else:
            folder_info = {"roomId": room_id, "path": folder_path}

        return self.__send_request__(endpoint='/rooms/folders/info/list', method='POST', headers=self.headers,
                                     json=folder_info).json()

    def send_email(self, email_addresses=None, subject=None, note=None, on_behalf_of=None, room_id=None,
                   document_id=None):

        email_info = {'recipients': email_addresses, 'subject': subject, 'note': note, 'onBehalfOf': on_behalf_of,
                      'readConfirmation': True}
        return self.__send_request__(
            endpoint='/rooms/{room_id}/documents/{document_id}/email/send'.format(room_id=room_id,
                                                                                  document_id=document_id),
            method='POST', headers=self.headers, json=email_info)

    def get_documents(self, room_id, folder_id=None, folder_path=None):

        document = namedtuple('Document', 'guid file_name sender date_of_submission modification_date file_type')

        if folder_id:
            folder_info = {"folderId": folder_id, "folders": False}
        elif folder_path:
            folder_info = {"folderPath": folder_path, "folders": False}
        else:
            folder_info = {"folders": False}
        try:
            document_objects = \
            self.__send_request__(endpoint='/rooms/{room_id}/documents/list'.format(room_id=room_id), method='POST',
                                  headers=self.headers, json=folder_info).json()['items']
            documents = [document(d['guid'], d['filename'], d['sender'], d['creationDate'], d['modifiedDate'],
                                  os.path.splitext(d['filename'])[1][1:].strip().lower()) for d in document_objects]
            return documents
        except Exception as e:
            print(e)
            return []

    def get_folders_and_documents(self, room_id=None, date_to_check=datetime(1900, 1, 1)):

        document = namedtuple('Document',
                              'guid file_name sender date_of_submission modification_date folder folder_id folder_location file_type')

        try:
            document_objects = \
            self.__send_request__(endpoint='/rooms/{room_id}/folders/documents/list'.format(room_id=room_id),
                                  method='POST', headers=self.headers).json()['documents']['items']
            documents = [
                document(d['guid'], d['filename'], d['sender'], d['creationDate'], d['modifiedDate'], d['folder'],
                         d['folderId'], d['folder'].rsplit('/', 1)[1],
                         os.path.splitext(d['filename'])[1][1:].strip().lower()) for d in document_objects if
                datetime.strptime(d['creationDate'].split('T')[0], '%Y-%m-%d') >= date_to_check
                or datetime.strptime(d['modifiedDate'].split('T')[0], '%Y-%m-%d') >= date_to_check]
            return documents
        except Exception as e:
            print(e)
            return []

    def get_entities(self, room_id=None):

        return self.__send_request__(endpoint='/rooms/{room_id}/entities'.format(room_id=room_id), method='GET',
                                     headers=self.headers)

    def get_entities_list(self, room_id=None):

        room_info = {"roomId": room_id, "fetchMembers": True}
        return self.__send_request__(endpoint='/rooms/entities', method='POST', headers=self.headers,
                                     json=room_info).json()

    def get_rooms(self):

        return self.__send_request__(endpoint='/rooms', method='GET', headers=self.headers)

    def download_original_document(self, document_id=None):

        return self.__send_request__(
            endpoint='/documents/{document_id}/download/original'.format(document_id=document_id), method='POST',
            headers=self.headers).content

    def download_protected_document(self, document_id=None):

        return self.__send_request__(endpoint='/documents/{document_id}/download'.format(document_id=document_id),
                                     method='POST', headers=self.headers).content

    def download_document(self, document_id=None):

        return self.__send_request__(endpoint='/documents/download', method='POST', headers=self.headers,
                                     json={"documentGuid": document_id}).content

    def get_document_info(self, document_id=None):

        return self.__send_request__(endpoint='/documents/{document_id}'.format(document_id=document_id), method='GET',
                                     headers=self.headers)['type'].lower()

    def get_document_activity(self, document_id=None, date_to_check=datetime(1900, 1, 1)):

        logs = []
        log = namedtuple('Log', 'user email time action')
        log_update_items = ['Updated via PC', 'Updated via Browser', 'Uploaded file via browser',
                            'updated file via a browser']
        activity_log = self.__send_request__(endpoint='/documents/activityLog', method='POST', headers=self.headers,
                                             json={"documentGuid": document_id}).json()
        print(activity_log)
        if activity_log['total'] > 0:
            logs = [log(a['user'], a['email'], a['time'], 'uploaded' if 'Uploaded' in a['activity'] else 'updated')
					for a in activity_log['items'] if a['time'] >= date_to_check
					and a['activity'] in log_update_items]
        return logs

    def create_document(self, room_guid=None, file_name=None, folder=None, folder_id=None, folder_guid=None):

        document_info = {"roomGuid": room_guid, "fileName": file_name, "folder": folder}
        return self.__send_request__(endpoint='/rooms/document/new/create', method='POST', headers=self.headers,
                                     json=document_info).json()

    def delete_documents(self, room_id=None, document_ids=None):

        document_info = {"documentGuids": document_ids}
        return self.__send_request__(endpoint='/rooms/{room_id}/documents/delete'.format(room_id=room_id),
                                     method='POST', headers=self.headers, json=document_info).json()

    def upload_document(self, room_id=None, file_path=None, folder_path=None, folder_id=None):

        files = {'data': open(file_path, 'rb')}
        file_guid = \
        self.__send_request__(endpoint='/rooms/{room_id}/documents/upload'.format(room_id=room_id), method='POST',
                              headers=self.auth_header, files=files).json()['guid']

        while True:

            file_state = self.__send_request__(
                endpoint='/rooms/{room_id}/documents/{document_id}'.format(room_id=room_id, document_id=file_guid),
                method='GET', headers=self.headers)['status']['documentState']
            if file_state == 'READY':
                if folder_id:
                    file_info = {'folderId': folder_id, 'documentGuids': [file_guid]}
                elif folder_path:
                    file_info = {'folder': folder_path, 'documentGuids': [file_guid]}
                file_guid = self.__send_request__(endpoint='/rooms/{room_id}/documents/submit'.format(room_id=room_id),
                                                  method='POST', headers=self.headers, json=file_info)
                break
            else:
                print(file_state)

    def folder_action(self, action=None, folder_path=None, folder_id=None, room_id=None, folder_name=None):

        if folder_path:
            main_folder_path = self.get_folder_info(room_id=room_id, folder_id=folder_id)['name']
            action_folder_path = folder_path if '_root' in main_folder_path else "{main_folder}/{folder_path}".format(
                main_folder=main_folder_path, folder_path=folder_path)
            folder_id = self.get_folder_info(room_id=room_id, folder_path=action_folder_path)['id']

        if action == 'Delete':
            delete_path = "{action_folder_path}/{folder_name}".format(action_folder_path=action_folder_path,
                                                                      folder_name=folder_name)
            self.delete_folders(room_id=room_id, folder_paths=[delete_path])
        elif action == 'Create':
            self.create_folder(room_id=room_id, folder_id=folder_id, new_folder_name=folder_name)
        else:
            pass

    def set_room_and_folder(self, folder_id=None, room_id=None):

        self.room_id = room_id
        self.base_folder_id = folder_id
        self.base_folder_name = self.get_folder_info(room_id=self.room_id, folder_id=self.base_folder_id)['name']
        self.folder_path = "{folder_type}".format(
            folder_type=folder_type) if "_root" in self.base_folder_name else "{folder_name}/{folder_type}".format(
            folder_name=self.base_folder_name, folder_type=folder_type)

        self.documents = self.get_documents(room_id=self.room_id, folder_path=self.folder_path)

    def revoke_permissions(self, user=None, room_id=None, folder_paths=None):

        permission_info = {"roomId": room_id,
                           "permittedEntitiesWithPermissions": [
                               {"permittedEntity": {"address": user, "entityType": "USER"},
                                "isDefault": False,
                                "revokePermissions": True,
                                "permissions": {
                                    "copy": False,
                                    "download": False,
                                    "downloadOriginal": False,
                                    "edit": False,
                                    "print": False,
                                    "progAccess": False,
                                    "spotlight": False,
                                    "watermark": True,
                                    "comment": False
                                },
                                "role": None
                                }],
                           "folderPathsOrIds": [{"path": f} for f in folder_paths]}
        return self.__send_request__(endpoint='/rooms/entities/permissions/change/bulk', method='POST',
                                     headers=self.headers, json=permission_info).json()

    def grant_full_permissions(self, user=None, room_id=None, folder_path=None):

        permission_info = {
            "roomId": room_id,
            "newPermissions": {
                "downloadOriginal": True,
                "download": True,
                "copy": True,
                "print": True,
                "edit": True,
                "spotlight": True,
                "progAccess": True,
                "watermark": True,
                "comment": True,
                "defaultExpirationDays": None,
                "expirationDate": "2021-05-26T23:59:59-0700"
            },
            "role": "CONTRIBUTORS",
            "folderPathOrId": {
                "path": folder_path
            },
            "addEntityToAllDocs": True,
            "isDefaultEntity": True,
            "roomEntities": [
                {
                    "address": user,
                    "entityType": "USER"
                },
            ],
            "isSendMail": False,
            "includeAllSubItems": False
        }

        return self.__send_request__(endpoint='/rooms/folders/entity/add', method='POST', headers=self.headers,
                                     json=permission_info).json()

    def set_admin_permissions(self, user=None, room_id=None):

        permission_info = {
            "roomId": room_id,
            "groupName": "Administrators",
            "membersList": [
                {
                    "entity": {
                        "address": user,
                        "entityType": "USER"
                    }
                }
            ],
            "managersList": []
        }
        print(permission_info)
        return self.__send_request__(endpoint='/rooms/group/members/add', method='POST', headers=self.headers,
                                     json=permission_info)

    def set_read_only_permissions(self, user=None, room_id=None, folder_path=None):

        permission_info = {
            "roomId": room_id,
            "newPermissions": {
                "downloadOriginal": True,
                "download": True,
                "copy": False,
                "print": True,
                "edit": False,
                "spotlight": False,
                "progAccess": False,
                "watermark": True,
                "comment": True,
                "defaultExpirationDays": None,
                "expirationDate": None
            },
            "role": "VISITORS",
            "folderPathOrId": {
                "path": folder_path
            },
            "addEntityToAllDocs": True,
            "isDefaultEntity": True,
            "roomEntities": [
                {
                    "address": user,
                    "entityType": "USER"
                },
            ],
            "isSendMail": False,
            "includeAllSubItems": False
        }

        return self.__send_request__(endpoint='/rooms/folders/entity/add', method='POST', headers=self.headers,
                                     json=permission_info).json()

    def session_logout(self):

        return self.__send_request__(endpoint='sessions/logout', method='POST')

    def create_users_and_groups_file(self):

        room = namedtuple('Room', 'id name')
        entity = namedtuple('Entity', 'name id role address')
        room_info = [room(r['id'], r['name']) for r in self.get_rooms()['items']]
        workbook = xlsxwriter.Workbook('watchdox_info.xlsx')
        worksheet = workbook.add_worksheet('Users and Groups')
        worksheet.set_column(0, 6, 35)
        bold_format = workbook.add_format({'bold': True})
        row = 1

        for room in room_info:
            col = 2
            worksheet.write(row, 0, room.name, bold_format)
            worksheet.write(row + 1, 1, 'Users')
            row += 2
            users = []
            groups = []
            entity_data = self.get_entities_list(room_id=room.id)
            for e in entity_data['items']:
                if e['entityType'] == "USER":
                    try:
                        users.append(entity(e['name'], e['id'], e['role'], e['address']))
                    except:
                        users.append(entity('No Name', e['id'], e['role'], e['address']))
                else:
                    g = {'name': e['name'], 'members': []}
                    try:
                        g['members'] = e['members']['userMembers']
                    except:
                        pass
                    groups.append(g)
            for u in users:
                worksheet.write(row, col, u.name)
                worksheet.write(row, col + 1, u.role)
                worksheet.write(row, col + 2, u.address)
                row += 1
            row += 1
            for g in groups:
                worksheet.write(row, 1, g['name'])
                for m in g['members']:
                    worksheet.write(row, col, m)
                    row += 1
            row += 2
        workbook.close()

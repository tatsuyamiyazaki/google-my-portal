function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SUNBIT - 統合ワークスペースポータル')
    .setFaviconUrl('https://img.icons8.com/fluency/96/hub.png');
}

function getUserInfo() {
  try {
    const user = Session.getActiveUser();
    const email = user.getEmail();
    
    return {
      email: email,
      name: email
    };
  } catch (error) {
    Logger.log('Error fetching user info: ' + error.toString());
    const user = Session.getActiveUser();
    return {
      email: user.getEmail(),
      name: user.getEmail()
    };
  }
}

function getGmailData() {
  try {
    const unreadThreads = GmailApp.search('is:unread in:inbox', 0, 100);
    const unreadCount = unreadThreads.length;
    let latestEmails = [];
    
    if (unreadCount > 0) {
      const threadsToShow = unreadThreads.slice(0, 100);
      
      for (let i = 0; i < threadsToShow.length; i++) {
        const thread = threadsToShow[i];
        const messages = thread.getMessages();
        const latestMessage = messages[messages.length - 1];
        
        let fromName = latestMessage.getFrom();
        const nameMatch = fromName.match(/^"?([^"<]+)"?\s*</);
        if (nameMatch) {
          fromName = nameMatch[1].trim();
        } else {
          const emailMatch = fromName.match(/<([^>]+)>/);
          if (emailMatch) {
            fromName = emailMatch[1];
          }
        }
        
        let snippet = '';
        try {
          const body = latestMessage.getPlainBody();
          snippet = body.substring(0, 100).replace(/\n/g, ' ').trim();
          if (body.length > 100) {
            snippet += '...';
          }
        } catch (e) {
          snippet = '本文を取得できませんでした';
        }
        
        latestEmails.push({
          subject: thread.getFirstMessageSubject() || '(件名なし)',
          from: fromName,
          date: latestMessage.getDate().getTime(),
          snippet: snippet,
          isUnread: true,
          messageCount: messages.length,
          threadId: thread.getId()
        });
      }
    }
    
    return {
      success: true,
      unreadCount: unreadCount,
      latestEmails: latestEmails,
      showingUnreadOnly: unreadCount > 0
    };
    
  } catch (error) {
    Logger.log('Error fetching Gmail data: ' + error.toString());
    return {
      success: false,
      unreadCount: 0,
      latestEmails: [],
      showingUnreadOnly: false,
      error: error.toString()
    };
  }
}

function markThreadAsRead(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    thread.markRead();
    
    return {
      success: true,
      message: 'メールを既読にしました'
    };
  } catch (error) {
    Logger.log('Error marking thread as read: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getEmailThread(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    const messages = thread.getMessages();
    
    const emailMessages = messages.map(function(message) {
      const attachments = message.getAttachments().map(function(attachment) {
        return {
          name: attachment.getName(),
          size: attachment.getSize(),
          contentType: attachment.getContentType()
        };
      });
      
      let fromDisplay = message.getFrom();
      const fromMatch = fromDisplay.match(/^"?([^"<]+)"?\s*<([^>]+)>$/);
      if (fromMatch) {
        fromDisplay = fromMatch[1].trim();
      } else if (fromDisplay.indexOf('<') === -1) {
        fromDisplay = fromDisplay.trim();
      } else {
        const emailOnlyMatch = fromDisplay.match(/<([^>]+)>/);
        if (emailOnlyMatch) {
          fromDisplay = emailOnlyMatch[1];
        }
      }
      
      let toDisplay = message.getTo();
      const toMatch = toDisplay.match(/^"?([^"<]+)"?\s*<([^>]+)>$/);
      if (toMatch) {
        toDisplay = toMatch[1].trim();
      } else if (toDisplay.indexOf('<') === -1) {
        toDisplay = toDisplay.trim();
      } else {
        const toEmailOnlyMatch = toDisplay.match(/<([^>]+)>/);
        if (toEmailOnlyMatch) {
          toDisplay = toEmailOnlyMatch[1];
        }
      }
      
      let replyTo = message.getFrom();
      const replyToMatch = replyTo.match(/<([^>]+)>/);
      if (replyToMatch) {
        replyTo = replyToMatch[1];
      }
      
      return {
        from: fromDisplay,
        to: toDisplay,
        replyTo: replyTo,
        subject: message.getSubject(),
        date: message.getDate().getTime(),
        body: message.getPlainBody(),
        isUnread: message.isUnread(),
        messageId: message.getId(),
        attachments: attachments
      };
    });
    
    emailMessages.reverse();
    
    return {
      success: true,
      subject: thread.getFirstMessageSubject(),
      messages: emailMessages,
      threadId: threadId
    };
    
  } catch (error) {
    Logger.log('Error fetching email thread: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function downloadAttachment(threadId, messageIndex, attachmentIndex) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    const messages = thread.getMessages();
    
    if (messageIndex >= messages.length) {
      return {
        success: false,
        error: 'メッセージが見つかりません'
      };
    }
    
    const message = messages[messageIndex];
    const attachments = message.getAttachments();
    
    if (attachmentIndex >= attachments.length) {
      return {
        success: false,
        error: '添付ファイルが見つかりません'
      };
    }
    
    const attachment = attachments[attachmentIndex];
    const blob = attachment.copyBlob();
    const data = Utilities.base64Encode(blob.getBytes());
    
    return {
      success: true,
      data: data,
      name: attachment.getName(),
      contentType: attachment.getContentType()
    };
    
  } catch (error) {
    Logger.log('Error downloading attachment: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getTodoTasks() {
  try {
    const taskLists = Tasks.Tasklists.list().items;
    if (!taskLists || taskLists.length === 0) {
      return {
        success: true,
        tasks: [],
        taskLists: []
      };
    }
    
    const allTasks = [];
    const listsInfo = [];
    
    for (let i = 0; i < taskLists.length; i++) {
      const taskList = taskLists[i];
      listsInfo.push({
        id: taskList.id,
        title: taskList.title
      });
      
      const tasks = Tasks.Tasks.list(taskList.id, {
        showCompleted: false,
        maxResults: 50
      });
      
      if (tasks.items) {
        tasks.items.forEach(function(task) {
          const subtasks = [];
          if (task.parent) {
            return;
          }
          
          if (tasks.items) {
            tasks.items.forEach(function(potentialChild) {
              if (potentialChild.parent === task.id) {
                subtasks.push({
                  id: potentialChild.id,
                  title: potentialChild.title,
                  completed: potentialChild.status === 'completed'
                });
              }
            });
          }
          
          allTasks.push({
            title: task.title,
            notes: task.notes || '',
            due: task.due || null,
            listName: taskList.title,
            id: task.id,
            listId: taskList.id,
            subtasks: subtasks
          });
        });
      }
    }
    
    allTasks.sort(function(a, b) {
      if (!a.due) return 1;
      if (!b.due) return -1;
      return new Date(a.due) - new Date(b.due);
    });
    
    return {
      success: true,
      tasks: allTasks,
      taskLists: listsInfo
    };
    
  } catch (error) {
    Logger.log('Error fetching tasks: ' + error.toString());
    return {
      success: false,
      tasks: [],
      taskLists: [],
      error: error.toString()
    };
  }
}

function addTodoTask(title, listId, dueDate, notes) {
  try {
    const taskLists = Tasks.Tasklists.list().items;
    let targetListId = listId;
    
    if (!targetListId && taskLists && taskLists.length > 0) {
      targetListId = taskLists[0].id;
    }
    
    const task = {
      title: title
    };
    
    if (dueDate) {
      task.due = new Date(dueDate).toISOString();
    }
    
    if (notes) {
      task.notes = notes;
    }
    
    const createdTask = Tasks.Tasks.insert(task, targetListId);
    
    return {
      success: true,
      message: 'タスクを追加しました',
      taskId: createdTask.id
    };
  } catch (error) {
    Logger.log('Error adding task: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function addSubtask(parentTaskId, listId, title) {
  try {
    const task = {
      title: title,
      parent: parentTaskId
    };
    
    Tasks.Tasks.insert(task, listId);
    
    return {
      success: true,
      message: 'サブタスクを追加しました'
    };
  } catch (error) {
    Logger.log('Error adding subtask: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateTodoTask(taskId, listId, title, dueDate, notes) {
  try {
    const task = Tasks.Tasks.get(listId, taskId);
    task.title = title;
    
    if (dueDate) {
      task.due = new Date(dueDate).toISOString();
    } else {
      task.due = null;
    }
    
    if (notes !== undefined) {
      task.notes = notes;
    }
    
    Tasks.Tasks.update(task, listId, taskId);
    
    return {
      success: true,
      message: 'タスクを更新しました'
    };
  } catch (error) {
    Logger.log('Error updating task: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function completeTodoTask(taskId, listId) {
  try {
    const task = Tasks.Tasks.get(listId, taskId);
    task.status = 'completed';
    Tasks.Tasks.update(task, listId, taskId);
    
    return {
      success: true,
      message: 'タスクを完了しました'
    };
  } catch (error) {
    Logger.log('Error completing task: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function deleteTodoTask(taskId, listId) {
  try {
    Tasks.Tasks.remove(listId, taskId);
    
    return {
      success: true,
      message: 'タスクを削除しました'
    };
  } catch (error) {
    Logger.log('Error deleting task: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getCalendarEvents(dateStr) {
  try {
    const date = new Date(dateStr + 'T00:00:00');
    const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
    const endOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    
    const calendars = CalendarApp.getAllCalendars();
    const allEvents = [];
    
    for (let i = 0; i < calendars.length; i++) {
      const calendar = calendars[i];
      const events = calendar.getEvents(startOfDay, endOfDay);
      
      events.forEach(function(event) {
        allEvents.push({
          id: event.getId(),
          title: event.getTitle(),
          startTime: event.getStartTime().getTime(),
          endTime: event.getEndTime().getTime(),
          location: event.getLocation() || '',
          description: event.getDescription() || '',
          calendarName: calendar.getName(),
          calendarId: calendar.getId(),
          isAllDay: event.isAllDayEvent(),
          canEdit: calendar.isOwnedByMe(),
          isTodo: false
        });
      });
    }
    
    try {
      const taskLists = Tasks.Tasklists.list().items;
      if (taskLists && taskLists.length > 0) {
        for (let i = 0; i < taskLists.length; i++) {
          const taskList = taskLists[i];
          const tasks = Tasks.Tasks.list(taskList.id, {
            showCompleted: false,
            maxResults: 100
          });
          
          if (tasks.items) {
            tasks.items.forEach(function(task) {
              if (task.due) {
                const taskDueDate = new Date(task.due);
                const taskDateOnly = new Date(taskDueDate.getFullYear(), taskDueDate.getMonth(), taskDueDate.getDate());
                const targetDateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
                
                if (taskDateOnly.getTime() === targetDateOnly.getTime()) {
                  allEvents.push({
                    id: 'todo_' + task.id,
                    title: '📋 ' + task.title,
                    startTime: taskDateOnly.getTime(),
                    endTime: taskDateOnly.getTime(),
                    location: '',
                    description: task.notes || '',
                    calendarName: 'Todo: ' + taskList.title,
                    calendarId: 'todo',
                    isAllDay: true,
                    canEdit: false,
                    isTodo: true,
                    taskId: task.id,
                    taskListId: taskList.id,
                    taskDue: task.due
                  });
                }
              }
            });
          }
        }
      }
    } catch (todoError) {
      Logger.log('Error fetching todos for calendar: ' + todoError.toString());
    }
    
    allEvents.sort(function(a, b) {
      return a.startTime - b.startTime;
    });
    
    return {
      success: true,
      events: allEvents,
      date: dateStr
    };
    
  } catch (error) {
    Logger.log('Error fetching calendar events: ' + error.toString());
    return {
      success: false,
      events: [],
      error: error.toString()
    };
  }
}

function addCalendarEvent(title, startTime, endTime, description, location) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    const event = calendar.createEvent(title, start, end, {
      description: description || '',
      location: location || ''
    });
    
    return {
      success: true,
      message: '予定を追加しました',
      eventId: event.getId()
    };
  } catch (error) {
    Logger.log('Error adding calendar event: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateCalendarEvent(eventId, title, startTime, endTime, description, location) {
  try {
    const event = CalendarApp.getEventById(eventId);
    
    if (!event) {
      return {
        success: false,
        error: 'イベントが見つかりません'
      };
    }
    
    event.setTitle(title);
    event.setTime(new Date(startTime), new Date(endTime));
    event.setDescription(description || '');
    event.setLocation(location || '');
    
    return {
      success: true,
      message: '予定を更新しました'
    };
  } catch (error) {
    Logger.log('Error updating calendar event: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function deleteCalendarEvent(eventId) {
  try {
    const event = CalendarApp.getEventById(eventId);
    
    if (!event) {
      return {
        success: false,
        error: 'イベントが見つかりません'
      };
    }
    
    event.deleteEvent();
    
    return {
      success: true,
      message: '予定を削除しました'
    };
  } catch (error) {
    Logger.log('Error deleting calendar event: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getMonthEvents(year, month) {
  try {
    const startOfMonth = new Date(year, month - 1, 1, 0, 0, 0);
    const endOfMonth = new Date(year, month, 0, 23, 59, 59);
    
    const calendars = CalendarApp.getAllCalendars();
    const eventDates = {};
    
    for (let i = 0; i < calendars.length; i++) {
      const calendar = calendars[i];
      const events = calendar.getEvents(startOfMonth, endOfMonth);
      
      events.forEach(function(event) {
        const eventDate = new Date(event.getStartTime());
        const dateKey = eventDate.getFullYear() + '-' + 
                       String(eventDate.getMonth() + 1).padStart(2, '0') + '-' + 
                       String(eventDate.getDate()).padStart(2, '0');
        
        if (!eventDates[dateKey]) {
          eventDates[dateKey] = 0;
        }
        eventDates[dateKey]++;
      });
    }
    
    try {
      const taskLists = Tasks.Tasklists.list().items;
      if (taskLists && taskLists.length > 0) {
        for (let i = 0; i < taskLists.length; i++) {
          const taskList = taskLists[i];
          const tasks = Tasks.Tasks.list(taskList.id, {
            showCompleted: false,
            maxResults: 100
          });
          
          if (tasks.items) {
            tasks.items.forEach(function(task) {
              if (task.due) {
                const taskDueDate = new Date(task.due);
                const dateKey = taskDueDate.getFullYear() + '-' + 
                               String(taskDueDate.getMonth() + 1).padStart(2, '0') + '-' + 
                               String(taskDueDate.getDate()).padStart(2, '0');
                
                if (taskDueDate >= startOfMonth && taskDueDate <= endOfMonth) {
                  if (!eventDates[dateKey]) {
                    eventDates[dateKey] = 0;
                  }
                  eventDates[dateKey]++;
                }
              }
            });
          }
        }
      }
    } catch (todoError) {
      Logger.log('Error fetching todos for month: ' + todoError.toString());
    }
    
    return {
      success: true,
      eventDates: eventDates
    };
    
  } catch (error) {
    Logger.log('Error fetching month events: ' + error.toString());
    return {
      success: false,
      eventDates: {},
      error: error.toString()
    };
  }
}

function getUserLinks() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    const foldersJson = userProperties.getProperty('customFolders');
    
    let links = linksJson ? JSON.parse(linksJson) : [];
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    
    return {
      success: true,
      links: links,
      folders: folders
    };
  } catch (error) {
    Logger.log('Error fetching user links: ' + error.toString());
    return {
      success: false,
      links: [],
      folders: [],
      error: error.toString()
    };
  }
}

function saveUserLink(title, url, icon, folderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    // URLからファビコンURLを生成
    let faviconUrl = '';
    try {
      const urlObj = new URL(url);
      faviconUrl = 'https://www.google.com/s2/favicons?domain=' + urlObj.hostname + '&sz=32';
    } catch (e) {
      faviconUrl = '';
    }
    
    const newLink = {
      id: Utilities.getUuid(),
      title: title,
      url: url,
      icon: icon || 'link',
      faviconUrl: faviconUrl,
      folderId: folderId || null,
      createdAt: new Date().getTime()
    };
    
    links.push(newLink);
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'リンクを追加しました',
      link: newLink
    };
  } catch (error) {
    Logger.log('Error saving user link: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateUserLink(linkId, title, url, icon, folderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    const linkIndex = links.findIndex(link => link.id === linkId);
    if (linkIndex === -1) {
      return {
        success: false,
        error: 'リンクが見つかりません'
      };
    }
    
    // URLからファビコンURLを生成
    let faviconUrl = '';
    try {
      const urlObj = new URL(url);
      faviconUrl = 'https://www.google.com/s2/favicons?domain=' + urlObj.hostname + '&sz=32';
    } catch (e) {
      faviconUrl = '';
    }
    
    links[linkIndex].title = title;
    links[linkIndex].url = url;
    links[linkIndex].icon = icon || 'link';
    links[linkIndex].faviconUrl = faviconUrl;
    links[linkIndex].folderId = folderId || null;
    links[linkIndex].updatedAt = new Date().getTime();
    
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'リンクを更新しました'
    };
  } catch (error) {
    Logger.log('Error updating user link: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function deleteUserLink(linkId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    links = links.filter(link => link.id !== linkId);
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'リンクを削除しました'
    };
  } catch (error) {
    Logger.log('Error deleting user link: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function saveUserFolder(name, color, parentFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const foldersJson = userProperties.getProperty('customFolders');
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    
    const newFolder = {
      id: Utilities.getUuid(),
      name: name,
      color: color || '#667eea',
      parentFolderId: parentFolderId || null,
      createdAt: new Date().getTime()
    };
    
    folders.push(newFolder);
    userProperties.setProperty('customFolders', JSON.stringify(folders));
    
    return {
      success: true,
      message: 'フォルダを追加しました',
      folder: newFolder
    };
  } catch (error) {
    Logger.log('Error saving user folder: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateUserFolder(folderId, name, color, parentFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const foldersJson = userProperties.getProperty('customFolders');
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    
    const folderIndex = folders.findIndex(folder => folder.id === folderId);
    if (folderIndex === -1) {
      return {
        success: false,
        error: 'フォルダが見つかりません'
      };
    }
    
    // 自分自身を親フォルダに設定しようとしていないかチェック
    if (parentFolderId === folderId) {
      return {
        success: false,
        error: 'フォルダを自分自身の中に配置することはできません'
      };
    }
    
    // 循環参照をチェック
    if (parentFolderId) {
      let currentParentId = parentFolderId;
      while (currentParentId) {
        if (currentParentId === folderId) {
          return {
            success: false,
            error: '循環参照が発生するため、このフォルダを親として設定できません'
          };
        }
        const parentFolder = folders.find(f => f.id === currentParentId);
        currentParentId = parentFolder ? parentFolder.parentFolderId : null;
      }
    }
    
    folders[folderIndex].name = name;
    folders[folderIndex].color = color || '#667eea';
    folders[folderIndex].parentFolderId = parentFolderId || null;
    folders[folderIndex].updatedAt = new Date().getTime();
    
    userProperties.setProperty('customFolders', JSON.stringify(folders));
    
    return {
      success: true,
      message: 'フォルダを更新しました'
    };
  } catch (error) {
    Logger.log('Error updating user folder: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function deleteUserFolder(folderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const foldersJson = userProperties.getProperty('customFolders');
    const linksJson = userProperties.getProperty('customLinks');
    
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    // このフォルダ内のリンクを「フォルダなし」に移動
    links = links.map(link => {
      if (link.folderId === folderId) {
        link.folderId = null;
      }
      return link;
    });
    
    // このフォルダの子フォルダの親を削除されたフォルダの親に変更
    const deletedFolder = folders.find(f => f.id === folderId);
    const newParentId = deletedFolder ? deletedFolder.parentFolderId : null;
    
    folders = folders.map(folder => {
      if (folder.parentFolderId === folderId) {
        folder.parentFolderId = newParentId;
      }
      return folder;
    });
    
    folders = folders.filter(folder => folder.id !== folderId);
    
    userProperties.setProperty('customFolders', JSON.stringify(folders));
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'フォルダを削除しました'
    };
  } catch (error) {
    Logger.log('Error deleting user folder: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 既存リンクにファビコンを追加する関数（オプション - 一度だけ実行）
function updateExistingLinksWithFavicons() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    links = links.map(function(link) {
      if (!link.faviconUrl) {
        try {
          const urlObj = new URL(link.url);
          link.faviconUrl = 'https://www.google.com/s2/favicons?domain=' + urlObj.hostname + '&sz=32';
        } catch (e) {
          link.faviconUrl = '';
        }
      }
      return link;
    });
    
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    Logger.log('既存リンクのファビコンを更新しました: ' + links.length + '件');
    
    return {
      success: true,
      message: '既存リンクのファビコンを更新しました',
      count: links.length
    };
  } catch (error) {
    Logger.log('Error updating existing links: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateLinkOrder(linkId, newIndex, newFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    const linkIndex = links.findIndex(link => link.id === linkId);
    if (linkIndex === -1) {
      return {
        success: false,
        error: 'リンクが見つかりません'
      };
    }
    
    // リンクを配列から取り出す
    const [movedLink] = links.splice(linkIndex, 1);
    
    // フォルダIDを更新
    movedLink.folderId = newFolderId || null;
    movedLink.updatedAt = new Date().getTime();
    
    // 新しい位置に挿入
    links.splice(newIndex, 0, movedLink);
    
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'リンクの順序を更新しました'
    };
  } catch (error) {
    Logger.log('Error updating link order: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateFolderOrder(folderId, newIndex, newParentFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const foldersJson = userProperties.getProperty('customFolders');
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    
    const folderIndex = folders.findIndex(folder => folder.id === folderId);
    if (folderIndex === -1) {
      return {
        success: false,
        error: 'フォルダが見つかりません'
      };
    }
    
    // 自分自身を親フォルダに設定しようとしていないかチェック
    if (newParentFolderId === folderId) {
      return {
        success: false,
        error: 'フォルダを自分自身の中に配置することはできません'
      };
    }
    
    // 循環参照をチェック
    if (newParentFolderId) {
      let currentParentId = newParentFolderId;
      let depth = 0;
      while (currentParentId && depth < 10) {
        if (currentParentId === folderId) {
          return {
            success: false,
            error: '循環参照が発生するため、このフォルダを親として設定できません'
          };
        }
        const parentFolder = folders.find(f => f.id === currentParentId);
        currentParentId = parentFolder ? parentFolder.parentFolderId : null;
        depth++;
      }
    }
    
    // フォルダを配列から取り出す
    const [movedFolder] = folders.splice(folderIndex, 1);
    
    // 親フォルダIDを更新
    movedFolder.parentFolderId = newParentFolderId || null;
    movedFolder.updatedAt = new Date().getTime();
    
    // 新しい位置に挿入
    folders.splice(newIndex, 0, movedFolder);
    
    userProperties.setProperty('customFolders', JSON.stringify(folders));
    
    return {
      success: true,
      message: 'フォルダの順序を更新しました'
    };
  } catch (error) {
    Logger.log('Error updating folder order: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function moveLink(linkId, direction, folderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    // 同じフォルダ内のリンクのみを取得
    const folderLinks = links.filter(link => (link.folderId || null) === (folderId || null));
    const otherLinks = links.filter(link => (link.folderId || null) !== (folderId || null));
    
    const currentIndex = folderLinks.findIndex(link => link.id === linkId);
    if (currentIndex === -1) {
      return { success: false, error: 'リンクが見つかりません' };
    }
    
    let newIndex = currentIndex;
    if (direction === 'up' && currentIndex > 0) {
      newIndex = currentIndex - 1;
    } else if (direction === 'down' && currentIndex < folderLinks.length - 1) {
      newIndex = currentIndex + 1;
    } else if (direction === 'top') {
      newIndex = 0;
    } else if (direction === 'bottom') {
      newIndex = folderLinks.length - 1;
    } else {
      return { success: true, message: '移動できません' };
    }
    
    // 配列内で要素を移動
    const [movedLink] = folderLinks.splice(currentIndex, 1);
    folderLinks.splice(newIndex, 0, movedLink);
    
    // 更新時刻を記録
    movedLink.updatedAt = new Date().getTime();
    
    // 全リンクを再構築
    const allLinks = [...otherLinks, ...folderLinks];
    
    userProperties.setProperty('customLinks', JSON.stringify(allLinks));
    
    return {
      success: true,
      message: 'リンクを移動しました'
    };
  } catch (error) {
    Logger.log('Error moving link: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function moveFolder(folderId, direction, parentFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const foldersJson = userProperties.getProperty('customFolders');
    let folders = foldersJson ? JSON.parse(foldersJson) : [];
    
    // 同じ親フォルダ内のフォルダのみを取得
    const siblingFolders = folders.filter(folder => (folder.parentFolderId || null) === (parentFolderId || null));
    const otherFolders = folders.filter(folder => (folder.parentFolderId || null) !== (parentFolderId || null));
    
    const currentIndex = siblingFolders.findIndex(folder => folder.id === folderId);
    if (currentIndex === -1) {
      return { success: false, error: 'フォルダが見つかりません' };
    }
    
    let newIndex = currentIndex;
    if (direction === 'up' && currentIndex > 0) {
      newIndex = currentIndex - 1;
    } else if (direction === 'down' && currentIndex < siblingFolders.length - 1) {
      newIndex = currentIndex + 1;
    } else if (direction === 'top') {
      newIndex = 0;
    } else if (direction === 'bottom') {
      newIndex = siblingFolders.length - 1;
    } else {
      return { success: true, message: '移動できません' };
    }
    
    // 配列内で要素を移動
    const [movedFolder] = siblingFolders.splice(currentIndex, 1);
    siblingFolders.splice(newIndex, 0, movedFolder);
    
    // 更新時刻を記録
    movedFolder.updatedAt = new Date().getTime();
    
    // 全フォルダを再構築
    const allFolders = [...otherFolders, ...siblingFolders];
    
    userProperties.setProperty('customFolders', JSON.stringify(allFolders));
    
    return {
      success: true,
      message: 'フォルダを移動しました'
    };
  } catch (error) {
    Logger.log('Error moving folder: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function moveLinkToFolder(linkId, targetFolderId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('customLinks');
    let links = linksJson ? JSON.parse(linksJson) : [];
    
    const linkIndex = links.findIndex(link => link.id === linkId);
    if (linkIndex === -1) {
      return { success: false, error: 'リンクが見つかりません' };
    }
    
    links[linkIndex].folderId = targetFolderId || null;
    links[linkIndex].updatedAt = new Date().getTime();
    
    userProperties.setProperty('customLinks', JSON.stringify(links));
    
    return {
      success: true,
      message: 'リンクを移動しました'
    };
  } catch (error) {
    Logger.log('Error moving link to folder: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Google Chatの未読スペース数を取得する
 */
function getChatData() {
  try {
    // 'state = "UNREAD"' フィルタを使用して、未読メッセージのあるスペースを取得
    const response = Chat.Spaces.list({
      'filter': 'state = "UNREAD"'
    });
    
    const unreadSpaces = response.spaces || [];
    const unreadCount = unreadSpaces.length;
    
    return {
      success: true,
      unreadCount: unreadCount
    };
    
  } catch (error) {
    Logger.log('Error fetching Google Chat data: ' + error.toString());
    return {
      success: false,
      unreadCount: 0,
      error: error.toString()
    };
  }
}

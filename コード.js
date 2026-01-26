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


function getRefreshIntervals() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const defaults = {
    emailRefreshInterval: 30000,
    chatRefreshInterval: 30000,
    calendarRefreshInterval: 60000,
    taskRefreshINterval: 60000
  };

  const toPositiveInt = function(value, fallback) {
    const num = Number(value);
    return Number.isFinite(num) && num > 0 ? Math.floor(num) : fallback;
  };

  const intervals = {
    emailRefreshInterval: toPositiveInt(scriptProperties.getProperty('emailRefreshInterval'), defaults.emailRefreshInterval),
    chatRefreshInterval: toPositiveInt(scriptProperties.getProperty('chatRefreshInterval'), defaults.chatRefreshInterval),
    calendarRefreshInterval: toPositiveInt(scriptProperties.getProperty('calendarRefreshInterval'), defaults.calendarRefreshInterval),
    taskRefreshINterval: toPositiveInt(scriptProperties.getProperty('taskRefreshINterval'), defaults.taskRefreshINterval)
  };

  return {
    success: true,
    intervals: intervals
  };
}

function getGmailData(listMode) {
  try {
    // 未読カウント機能は無効化されました
    // const unreadThreads = GmailApp.search('is:unread in:inbox', 0, 100);
    const unreadCount = 0; // 常に0を返す
    const normalizedMode = listMode === 'inbox' ? 'inbox' : 'unread';
    const listQuery = normalizedMode === 'inbox' ? 'in:inbox' : 'is:unread in:inbox';
    const listThreads = GmailApp.search(listQuery, 0, 100);
    let latestEmails = [];
    
    if (listThreads.length > 0) {
      const threadsToShow = listThreads.slice(0, 100);
      
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
          isUnread: normalizedMode === 'unread' ? true : thread.isUnread(),
          messageCount: messages.length,
          threadId: thread.getId()
        });
      }
    }
    
    return {
      success: true,
      unreadCount: 0, // 未読カウント機能は無効化
      latestEmails: latestEmails,
      listMode: normalizedMode
    };
    
  } catch (error) {
    Logger.log('Error fetching Gmail data: ' + error.toString());
    return {
      success: false,
      unreadCount: 0,
      latestEmails: [],
      listMode: listMode === 'inbox' ? 'inbox' : 'unread',
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

function deleteEmailThread(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    thread.moveToTrash();
    return {
      success: true,
      message: 'メールをゴミ箱に移動しました'
    };
  } catch (error) {
    Logger.log('Error deleting thread: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function archiveEmailThread(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    thread.moveToArchive();
    return {
      success: true,
      message: 'メールをアーカイブしました'
    };
  } catch (error) {
    Logger.log('Error archiving thread: ' + error.toString());
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
      let allAttachments = [];
      try {
        allAttachments = message.getAttachments();
      } catch (attachError) {
        Logger.log('Error getting attachments: ' + attachError.toString());
      }

      const attachments = [];
      const inlineImages = {};

      allAttachments.forEach(function(attachment) {
        try {
          const contentType = attachment.getContentType() || '';
          const isImage = contentType.indexOf('image/') === 0;
          const name = attachment.getName() || '';

          // Check if it's an inline image
          let cid = null;
          try {
            cid = attachment.getCid();
          } catch (cidError) {
            // getCid() might throw an error for some attachments
          }

          const isInline = attachment.isInlineImage() || (isImage && cid);

          if (isInline && isImage) {
            // Limit inline image size to prevent memory issues (5MB max)
            const size = attachment.getSize();
            if (size < 5 * 1024 * 1024) {
              const data = Utilities.base64Encode(attachment.getBytes());
              const dataUri = 'data:' + contentType + ';base64,' + data;

              if (cid) {
                // Store with original CID
                inlineImages[cid] = dataUri;
                // Also store without angle brackets if CID has them
                const cleanCid = cid.replace(/^<|>$/g, '');
                if (cleanCid !== cid) {
                  inlineImages[cleanCid] = dataUri;
                }
              }

              // Also try to map by filename for some email clients
              if (name) {
                inlineImages[name] = dataUri;
              }
            }
          } else if (!isInline || !isImage) {
            attachments.push({
              name: name || 'attachment',
              size: attachment.getSize() || 0,
              contentType: contentType
            });
          }
        } catch (attError) {
          Logger.log('Error processing attachment: ' + attError.toString());
          // Continue processing other attachments
        }
      });

      let body = '';
      try {
        body = message.getBody();
        // Check if body is plain text (no significant HTML structure)
        const isPlainText = !/<(html|body|div|p|span|table|br)\b/i.test(body);
        if (isPlainText && body.trim()) {
          // Plain text email - escape special characters and convert line breaks
          body = body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
          // Convert URLs to clickable links
          body = body.replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank">$1</a>');
          body = body.replace(/\n/g, '<br>');
        }
      } catch (bodyError) {
        Logger.log('Error getting HTML body: ' + bodyError.toString());
        try {
          body = message.getPlainBody();
          // Escape and convert plain text to HTML
          body = body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
          // Convert URLs to clickable links
          body = body.replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank">$1</a>');
          body = body.replace(/\n/g, '<br>');
        } catch (plainError) {
          body = '<p>本文を取得できませんでした</p>';
        }
      }

      // Replace CID references with base64 data URIs for inline images
      for (const cid in inlineImages) {
        try {
          // Escape special regex characters in CID
          const escapedCid = cid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          // Handle various formats: cid:xxx, CID:xxx, cid:<xxx>, etc.
          body = body.replace(new RegExp('cid:' + escapedCid, 'gi'), inlineImages[cid]);
          // Also handle with angle brackets
          body = body.replace(new RegExp('cid:<' + escapedCid + '>', 'gi'), inlineImages[cid]);
          // Handle src attribute with filename directly (some clients)
          body = body.replace(new RegExp('src=["\']' + escapedCid + '["\']', 'gi'), 'src="' + inlineImages[cid] + '"');
        } catch (replaceError) {
          Logger.log('Error replacing CID: ' + replaceError.toString());
        }
      }
      
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
        body: body, // Use the modified body with inline images
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

function getAccessibleCalendars() {
  // Prefer the advanced Calendar API so we can honor access roles and fall back to CalendarApp if unavailable
  try {
    const calendars = [];
    let pageToken = null;
    
    do {
      const response = Calendar.CalendarList.list({
        maxResults: 250,
        minAccessRole: 'reader',
        pageToken: pageToken
      });
      
      if (response.items) {
        response.items.forEach(function(item) {
          calendars.push({
            id: item.id,
            name: item.summary,
            canEdit: item.accessRole === 'owner' || item.accessRole === 'writer',
            source: 'api'
          });
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    if (calendars.length > 0) {
      return calendars;
    }
  } catch (apiError) {
    Logger.log('Advanced Calendar API unavailable, falling back to CalendarApp: ' + apiError.toString());
  }
  
  try {
    return CalendarApp.getAllCalendars().map(function(cal) {
      return {
        id: cal.getId(),
        name: cal.getName(),
        canEdit: cal.isOwnedByMe(),
        source: 'app',
        appCalendar: cal
      };
    });
  } catch (fallbackError) {
    Logger.log('CalendarApp.getAllCalendars failed: ' + fallbackError.toString());
    return [];
  }
}

function getCalendarEvents(dateStr) {
  try {
    const date = new Date(dateStr + 'T00:00:00');
    const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
    const endOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    const allEvents = [];
    const calendars = getAccessibleCalendars();
    
    calendars.forEach(function(calendar) {
      try {
        if (calendar.source === 'api') {
          let pageToken = null;
          do {
            const response = Calendar.Events.list(calendar.id, {
              timeMin: startOfDay.toISOString(),
              timeMax: endOfDay.toISOString(),
              singleEvents: true,
              orderBy: 'startTime',
              maxResults: 2500,
              pageToken: pageToken
            });
            
            const events = response.items || [];
            events.forEach(function(event) {
              const start = event.start.dateTime ? new Date(event.start.dateTime) : new Date(event.start.date);
              const end = event.end.dateTime ? new Date(event.end.dateTime) : new Date(event.end.date);

              // 会議リンクの取得
              let conferenceLink = '';
              let conferenceType = '';

              // hangoutLink (Google Meet)
              if (event.hangoutLink) {
                conferenceLink = event.hangoutLink;
                conferenceType = 'meet';
              }

              // conferenceData (より詳細な会議情報)
              if (event.conferenceData && event.conferenceData.entryPoints) {
                const videoEntry = event.conferenceData.entryPoints.find(function(ep) {
                  return ep.entryPointType === 'video';
                });
                if (videoEntry && videoEntry.uri) {
                  conferenceLink = videoEntry.uri;
                  // 会議タイプの判定
                  if (videoEntry.uri.indexOf('meet.google.com') !== -1) {
                    conferenceType = 'meet';
                  } else if (videoEntry.uri.indexOf('teams.microsoft.com') !== -1 || videoEntry.uri.indexOf('teams.live.com') !== -1) {
                    conferenceType = 'teams';
                  } else if (videoEntry.uri.indexOf('zoom.us') !== -1) {
                    conferenceType = 'zoom';
                  } else {
                    conferenceType = 'other';
                  }
                }
              }

              // locationからも会議リンクを検出
              if (!conferenceLink && event.location) {
                if (event.location.indexOf('meet.google.com') !== -1) {
                  conferenceLink = event.location.match(/https?:\/\/meet\.google\.com\/[^\s]*/)?.[0] || '';
                  conferenceType = 'meet';
                } else if (event.location.indexOf('teams.microsoft.com') !== -1 || event.location.indexOf('teams.live.com') !== -1) {
                  conferenceLink = event.location.match(/https?:\/\/teams\.(microsoft|live)\.com\/[^\s]*/)?.[0] || '';
                  conferenceType = 'teams';
                } else if (event.location.indexOf('zoom.us') !== -1) {
                  conferenceLink = event.location.match(/https?:\/\/[^\s]*zoom\.us\/[^\s]*/)?.[0] || '';
                  conferenceType = 'zoom';
                }
              }

              // 繰り返し予定のルール取得
              let recurrenceRule = '';
              const isRecurring = !!event.recurringEventId || (event.recurrence && event.recurrence.length > 0);

              if (event.recurrence && event.recurrence.length > 0) {
                // 元のイベント自体に繰り返しルールがある場合
                recurrenceRule = event.recurrence[0];
              } else if (event.recurringEventId) {
                // 展開されたインスタンスの場合、元のイベントから取得を試みる
                try {
                  const originalEvent = Calendar.Events.get(calendar.id, event.recurringEventId);
                  if (originalEvent.recurrence && originalEvent.recurrence.length > 0) {
                    recurrenceRule = originalEvent.recurrence[0];
                  }
                } catch (e) {
                  // 元イベントの取得に失敗した場合は空のまま
                  Logger.log('Could not fetch original recurring event: ' + e.toString());
                }
              }

              allEvents.push({
                id: event.id,
                title: event.summary || '(無題)',
                startTime: start.getTime(),
                endTime: end.getTime(),
                location: event.location || '',
                description: event.description || '',
                calendarName: calendar.name,
                calendarId: calendar.id,
                isAllDay: !event.start.dateTime,
                isRecurring: isRecurring,
                recurrenceRule: recurrenceRule,
                conferenceLink: conferenceLink,
                conferenceType: conferenceType,
                canEdit: calendar.canEdit,
                isTodo: false,
                colorId: event.colorId || ''
              });
            });
            
            pageToken = response.nextPageToken;
          } while (pageToken);
        } else if (calendar.appCalendar) {
          const events = calendar.appCalendar.getEvents(startOfDay, endOfDay);

          events.forEach(function(event) {
            const location = event.getLocation() || '';
            const description = event.getDescription() || '';

            // CalendarAppからは会議リンクを直接取得できないので、locationとdescriptionから検出
            let conferenceLink = '';
            let conferenceType = '';

            const textToSearch = location + ' ' + description;
            if (textToSearch.indexOf('meet.google.com') !== -1) {
              conferenceLink = textToSearch.match(/https?:\/\/meet\.google\.com\/[^\s<]*/)?.[0] || '';
              conferenceType = 'meet';
            } else if (textToSearch.indexOf('teams.microsoft.com') !== -1 || textToSearch.indexOf('teams.live.com') !== -1) {
              conferenceLink = textToSearch.match(/https?:\/\/teams\.(microsoft|live)\.com\/[^\s<]*/)?.[0] || '';
              conferenceType = 'teams';
            } else if (textToSearch.indexOf('zoom.us') !== -1) {
              conferenceLink = textToSearch.match(/https?:\/\/[^\s<]*zoom\.us\/[^\s<]*/)?.[0] || '';
              conferenceType = 'zoom';
            }

            // CalendarApp からはcolorIdを直接取得できないため空にする
            allEvents.push({
              id: event.getId(),
              title: event.getTitle(),
              startTime: event.getStartTime().getTime(),
              endTime: event.getEndTime().getTime(),
              location: location,
              description: description,
              calendarName: calendar.name,
              calendarId: calendar.id,
              isAllDay: event.isAllDayEvent(),
              isRecurring: event.isRecurring(),
              recurrenceRule: '',
              conferenceLink: conferenceLink,
              conferenceType: conferenceType,
              canEdit: calendar.canEdit,
              isTodo: false,
              colorId: ''
            });
          });
        }
      } catch (calendarError) {
        Logger.log('Error fetching events for calendar ' + calendar.id + ': ' + calendarError.toString());
      }
    });
    
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

function addCalendarEvent(title, startTime, endTime, description, location, recurrenceRule, colorId) {
  try {
    const start = new Date(startTime);
    const end = new Date(endTime);
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const calendarId = CalendarApp.getDefaultCalendar().getId();

    // Calendar APIを使用して作成（colorIdや繰り返しルールに対応）
    const body = {
      summary: title,
      description: description || '',
      location: location || '',
      start: { dateTime: start.toISOString(), timeZone: tz },
      end: { dateTime: end.toISOString(), timeZone: tz }
    };

    // 繰り返しルールがある場合
    if (recurrenceRule) {
      body.recurrence = [recurrenceRule];
    }

    // 色の設定（1-11のcolorId）
    if (colorId && colorId !== '') {
      body.colorId = colorId;
    }

    const createdEvent = Calendar.Events.insert(body, calendarId);

    return {
      success: true,
      message: recurrenceRule ? '繰り返し予定を追加しました' : '予定を追加しました',
      eventId: createdEvent.id
    };
  } catch (error) {
    Logger.log('Error adding calendar event: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function updateCalendarEvent(eventId, calendarId, title, startTime, endTime, description, location, isAllDay, recurrenceRule, colorId) {
  try {
    // Prefer Calendar API when calendarId is provided
    if (calendarId) {
      try {
        const startDate = new Date(startTime);
        const endDate = new Date(endTime);
        const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';

        // まず現在のイベントを取得して繰り返し予定かチェック
        const currentEvent = Calendar.Events.get(calendarId, eventId);
        const isRecurringInstance = !!currentEvent.recurringEventId;
        const targetEventId = isRecurringInstance ? currentEvent.recurringEventId : eventId;

        const body = {
          summary: title,
          description: description || '',
          location: location || ''
        };

        if (isAllDay) {
          body.start = { date: Utilities.formatDate(startDate, tz, 'yyyy-MM-dd') };
          body.end = { date: Utilities.formatDate(endDate, tz, 'yyyy-MM-dd') };
        } else {
          body.start = { dateTime: startDate.toISOString(), timeZone: tz };
          body.end = { dateTime: endDate.toISOString(), timeZone: tz };
        }

        // 繰り返しルールの設定
        if (recurrenceRule) {
          body.recurrence = [recurrenceRule];
        } else if (isRecurringInstance || (currentEvent.recurrence && currentEvent.recurrence.length > 0)) {
          // 繰り返しを解除する場合は空配列
          body.recurrence = [];
        }

        // 色の設定（1-11のcolorId、空文字の場合は色をリセット）
        if (colorId !== undefined) {
          if (colorId && colorId !== '') {
            body.colorId = colorId;
          } else {
            // 色をリセットする場合はnullを設定
            body.colorId = null;
          }
        }

        // 繰り返し予定のインスタンスの場合は元のイベントを更新
        if (isRecurringInstance) {
          // 繰り返し予定全体を更新
          Calendar.Events.patch(body, calendarId, targetEventId);
        } else {
          Calendar.Events.patch(body, calendarId, eventId);
        }

        return {
          success: true,
          message: '予定を更新しました'
        };
      } catch (apiError) {
        Logger.log('Calendar API update failed: ' + apiError.toString());
        // Fall through to CalendarApp fallback
      }
    }

    const event = CalendarApp.getEventById(eventId);

    if (!event) {
      return {
        success: false,
        error: 'イベントが見つかりません'
      };
    }

    event.setTitle(title);
    event.setDescription(description || '');
    event.setLocation(location || '');

    if (isAllDay) {
      const startDate = new Date(startTime);
      const endDate = new Date(endTime);
      endDate.setDate(endDate.getDate() - 1);
      event.setAllDayDates(startDate, endDate);
    } else {
      event.setTime(new Date(startTime), new Date(endTime));
    }

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

function deleteCalendarEvent(eventId, calendarId) {
  try {
    if (calendarId) {
      try {
        Calendar.Events.remove(calendarId, eventId);
        return {
          success: true,
          message: '予定を削除しました'
        };
      } catch (apiError) {
        Logger.log('Calendar API delete failed: ' + apiError.toString());
        // Fall back to CalendarApp
      }
    }
    
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
    
    const eventDates = {};
    const calendars = getAccessibleCalendars();
    
    calendars.forEach(function(calendar) {
      try {
        if (calendar.source === 'api') {
          let pageToken = null;
          do {
            const response = Calendar.Events.list(calendar.id, {
              timeMin: startOfMonth.toISOString(),
              timeMax: endOfMonth.toISOString(),
              singleEvents: true,
              orderBy: 'startTime',
              maxResults: 2500,
              pageToken: pageToken
            });
            
            const events = response.items || [];
            events.forEach(function(event) {
              const eventDate = event.start.dateTime ? new Date(event.start.dateTime) : new Date(event.start.date);
              const dateKey = eventDate.getFullYear() + '-' + 
                             String(eventDate.getMonth() + 1).padStart(2, '0') + '-' + 
                             String(eventDate.getDate()).padStart(2, '0');
              
              eventDates[dateKey] = (eventDates[dateKey] || 0) + 1;
            });
            
            pageToken = response.nextPageToken;
          } while (pageToken);
        } else if (calendar.appCalendar) {
          const events = calendar.appCalendar.getEvents(startOfMonth, endOfMonth);
          
          events.forEach(function(event) {
            const eventDate = new Date(event.getStartTime());
            const dateKey = eventDate.getFullYear() + '-' + 
                           String(eventDate.getMonth() + 1).padStart(2, '0') + '-' + 
                           String(eventDate.getDate()).padStart(2, '0');
            
            eventDates[dateKey] = (eventDates[dateKey] || 0) + 1;
          });
        }
      } catch (calendarError) {
        Logger.log('Error fetching month events for calendar ' + calendar.id + ': ' + calendarError.toString());
      }
    });
    
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

// --- Memo Functions ---

function getMemoSheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('work_management_sheet_id');
  const sheetName = scriptProperties.getProperty('memo_sheet_name');
  const ss = SpreadsheetApp.openById(sheetId);
  return ss.getSheetByName(sheetName);
}

function getMemoColumnIndices() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    memo: columnLetterToIndex(scriptProperties.getProperty('memo_col')),
    tag: columnLetterToIndex(scriptProperties.getProperty('tag_col')),
    color: columnLetterToIndex(scriptProperties.getProperty('color_col'))
  };
}

function columnLetterToIndex(letter) {
  if (!letter) return 1;
  letter = letter.toUpperCase();
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

function getMemos() {
  try {
    const sheet = getMemoSheet();
    const cols = getMemoColumnIndices();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return { success: true, memos: [] };
    }

    const memos = [];
    const maxCol = Math.max(cols.memo, cols.tag, cols.color);
    const data = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

    for (let i = 0; i < data.length; i++) {
      const content = data[i][cols.memo - 1];
      if (!content) continue;

      const tagsStr = data[i][cols.tag - 1] || '';
      const tags = tagsStr ? tagsStr.split(',').map(t => t.trim()).filter(t => t) : [];
      const color = data[i][cols.color - 1] || '#667eea';

      memos.push({
        id: i + 2, // 行番号をIDとして使用（2行目から開始）
        content: content,
        tags: tags,
        color: color,
        createdAt: new Date().getTime(),
        updatedAt: new Date().getTime()
      });
    }

    return { success: true, memos: memos.reverse() }; // 新しいものが上に来るよう逆順
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function saveMemo(content, tags, color) {
  try {
    const sheet = getMemoSheet();
    const cols = getMemoColumnIndices();
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    const tagsStr = (tags || []).join(', ');
    const memoColor = color || '#667eea';

    sheet.getRange(newRow, cols.memo).setValue(content);
    sheet.getRange(newRow, cols.tag).setValue(tagsStr);
    sheet.getRange(newRow, cols.color).setValue(memoColor);

    const newMemo = {
      id: newRow,
      content: content,
      tags: tags || [],
      color: memoColor,
      createdAt: new Date().getTime(),
      updatedAt: new Date().getTime()
    };

    return { success: true, memo: newMemo };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function updateMemo(id, content, tags, color) {
  try {
    const sheet = getMemoSheet();
    const cols = getMemoColumnIndices();
    const row = parseInt(id);

    if (row < 2 || row > sheet.getLastRow()) {
      return { success: false, error: 'Memo not found' };
    }

    sheet.getRange(row, cols.memo).setValue(content);
    if (tags) {
      const tagsStr = tags.join(', ');
      sheet.getRange(row, cols.tag).setValue(tagsStr);
    }
    if (color) {
      sheet.getRange(row, cols.color).setValue(color);
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function deleteMemo(id) {
  try {
    const sheet = getMemoSheet();
    const row = parseInt(id);

    if (row < 2 || row > sheet.getLastRow()) {
      return { success: false, error: 'Memo not found' };
    }

    sheet.deleteRow(row);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// --- Fixed Link Functions ---

function getFixedLinks() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const linksJson = userProperties.getProperty('fixedLinks');
    
    if (!linksJson) {
      // Return defaults if not set
      const defaultLinks = [
        { id: 'gmail', name: 'Gmail', url: 'https://mail.google.com', icon: 'email', color: 'linear-gradient(135deg, #F093FB 0%, #F5576C 100%)' },
        { id: 'chat', name: 'Chat', url: 'https://chat.google.com', icon: 'chat', color: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)' },
        { id: 'calendar', name: 'Calendar', url: 'https://calendar.google.com', icon: 'event', color: 'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)' },
        { id: 'drive', name: 'Drive', url: 'https://drive.google.com', icon: 'cloud', color: 'linear-gradient(135deg, #30cfd0 0%, #330867 100%)' },
        { id: 'salesforce', name: 'Salesforce', url: 'https://sunbit.my.salesforce.com', icon: 'cloud', color: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)' },
        { id: 'gemini', name: 'Gemini', url: 'https://gemini.google.com', icon: 'psychology', color: 'linear-gradient(135deg, #8E54E9 0%, #4776E6 100%)' },
        { id: 'genspark', name: 'Genspark', url: 'https://www.genspark.ai', icon: 'psychology', color: 'linear-gradient(135deg, #8E54E9 0%, #4776E6 100%)' },
        { id: 'notebooklm', name: 'NotebookLM', url: 'https://notebooklm.google.com', icon: 'psychology', color: 'linear-gradient(135deg, #8E54E9 0%, #4776E6 100%)' },
        { id: 'workspace_studio', name: 'Workspace Studio', url: 'https://studio.google.com', icon: 'build', color: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' },
        { id: 'redmine', name: 'Redmine', url: '#', icon: 'list', color: 'linear-gradient(135deg, #ff9a9e 0%, #fecfef 99%, #fecfef 100%)' }
      ];
      return { success: true, links: defaultLinks };
    }
    
    return { success: true, links: JSON.parse(linksJson) };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function saveFixedLinks(links) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('fixedLinks', JSON.stringify(links));
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// --- Department Metrics Functions ---

function getDepartmentMetrics() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // スプレッドシートのURLを取得
    const sheetUrl = scriptProperties.getProperty('management_index_sheet_url');
    if (!sheetUrl) {
      return { success: false, error: 'スプレッドシートURLが設定されていません' };
    }
    
    // スプレッドシートを開く
    const ss = SpreadsheetApp.openByUrl(sheetUrl);
    const sheetName = scriptProperties.getProperty('dashboard_sheet_name') || 'Dashboard';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, error: 'シート "' + sheetName + '" が見つかりません' };
    }
    
    // プロパティから設定値を取得
    const e1Row = parseInt(scriptProperties.getProperty('e1') || '2');
    const e2Row = parseInt(scriptProperties.getProperty('e2') || '3');
    const ebRow = parseInt(scriptProperties.getProperty('eb') || '4');
    
    const salesCol = scriptProperties.getProperty('sales') || 'B';
    const salesPlanCol = scriptProperties.getProperty('sales_plan') || 'C';
    const salesRateCol = scriptProperties.getProperty('sales_rate') || 'D';
    const profitCol = scriptProperties.getProperty('profit') || 'E';
    const profitPlanCol = scriptProperties.getProperty('profit_plan') || 'F';
    const profitRateCol = scriptProperties.getProperty('profit_rate') || 'G';
    const profitPerTimeCol = scriptProperties.getProperty('profit_per_time') || 'H';
    const profitPerTimePlanCol = scriptProperties.getProperty('profit_per_time_plan') || 'I';
    const profitPerTimeRateCol = scriptProperties.getProperty('profit_per_time_rate') || 'J';
    
    // データを取得する関数
    function getDeptData(row) {
      return {
        sales: sheet.getRange(salesCol + row).getValue() || '',
        salesPlan: sheet.getRange(salesPlanCol + row).getValue() || '',
        salesRate: sheet.getRange(salesRateCol + row).getValue() || '',
        profit: sheet.getRange(profitCol + row).getValue() || '',
        profitPlan: sheet.getRange(profitPlanCol + row).getValue() || '',
        profitRate: sheet.getRange(profitRateCol + row).getValue() || '',
        profitPerTime: sheet.getRange(profitPerTimeCol + row).getValue() || '',
        profitPerTimePlan: sheet.getRange(profitPerTimePlanCol + row).getValue() || '',
        profitPerTimeRate: sheet.getRange(profitPerTimeRateCol + row).getValue() || ''
      };
    }
    
    const metrics = {
      E1: getDeptData(e1Row),
      E2: getDeptData(e2Row),
      EB: getDeptData(ebRow)
    };
    
    return {
      success: true,
      metrics: metrics
    };
  } catch (error) {
    Logger.log('Error fetching department metrics: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

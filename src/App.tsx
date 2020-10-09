import React from 'react';
import './App.css';
import { Login, PeoplePicker, Person } from '@microsoft/mgt-react';
import { PersonViewType, PersonCardInteraction} from '@microsoft/mgt';
import { useGet } from './mgt';
import { DefaultButton, DetailsList, IColumn, PrimaryButton, SelectionMode, Spinner, SpinnerSize, Stack, TextField } from '@fluentui/react';
import { Message } from '@microsoft/microsoft-graph-types'

function App() {
  return (
    <div className="App">
      <header>
        <Login></Login>
      </header>
      <div className="Content">
        <Mail></Mail>
      </div>
    </div>
  );
}

function Mail() {

  let [messages, messagesLoading] = useGet('/me/messages');

  if (messagesLoading) {
    return <Spinner size={SpinnerSize.large} label="loading messages"></Spinner>
  }

  if (messages && messages.value && messages.value.length) {
    const items = messages.value.map((m: Message) => {
      return {
        key: m.id,
        from: m.sender?.emailAddress?.address,
        subject: m.subject,
        preview: m.bodyPreview
      }
    })

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'From',
        minWidth: 150,
        maxWidth: 160,
        onRender: (item) => <Person personQuery={item.from} 
          view={PersonViewType.oneline} 
          fetchImage 
          personCardInteraction={PersonCardInteraction.hover}>
        </Person>
      },
      {
        key: 'column2',
        name: 'Subject',
        minWidth: 100,
        maxWidth: 200,
        fieldName: 'subject' 
      },
      {
        key: 'column3',
        name: 'Body',
        minWidth: 100,
        fieldName: 'preview' 
      }
    ] 

    return <div>
      <div className="MessagesHeader">
        {/* <h2>Messages ({messages.value.length})</h2> */}
        <DefaultButton text="New"></DefaultButton>
      </div>
      <div className="MessagesMain">
        <div className="MessagesList">
          <DetailsList selectionMode={SelectionMode.none} items={items} columns={columns} ></DetailsList>
        </div>
        <div className="MessagesBody">
          <NewMessage></NewMessage>
        </div>
      </div>
    </div>
  }

  return <div></div>
}

function NewMessage() {

  return <div>
    <Stack>
      <PeoplePicker placeholder="To"></PeoplePicker>
      <PeoplePicker placeholder="Cc"></PeoplePicker>
      <TextField placeholder="Add a subject"></TextField>
      <TextField multiline={true}></TextField>
      <PrimaryButton text="Send"></PrimaryButton>
    </Stack>

  </div>
}

export default App;

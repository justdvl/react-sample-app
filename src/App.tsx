import React from 'react';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { mergeStyleSets } from 'office-ui-fabric-react';
import connection from './eWayAPI/Connector';
import { TContactsResopnse } from './eWayAPI/ContactsResponse';
import { SearchBox, ISearchBoxStyles } from 'office-ui-fabric-react/lib/SearchBox';

const css = mergeStyleSets({
    loadingDiv: {
        width: '50vw',
        position: 'absolute',
        left: '25vw',
        top: '40vh'
    }
})


const searchBoxStyles: Partial<ISearchBoxStyles> = { 
    root: {
        margin: 20,
        width: 380,
        border: "4px solid purple",
        selectors: {
            "[HighContrastSelector]": {
            borderColor: "WindowText"
            },
            ":hover": {
                borderColor: "${theme.brandDeepSkyBlue}",
                selectors: {
                    "[HighContrastSelector]": {
                        borderColor: "Highlight"
                    }
                }
            }
        }
    }
};


// This is a React Hook component.
function App() {

    const [found, setFound] = React.useState<boolean>(true);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [contact, setContact] = React.useState<Object | any>({});

    const onSearch = (query: string) => {

        setFound(true)
        if (!/.+@.+\.[A-Za-z]+$/.test(query)) { /* return true */ 
           window.alert("You must input valid email address")
           return
        }

        setLoading(true); 

        try{
            connection.callMethod(
                'SearchContacts',
                {
                    transmitObject: {
                        Email1Address: query
                    },
                    includeProfilePictures: true
                },
                (result: TContactsResopnse) => {
                    console.log("result", result)
                    setLoading(false)
                    if (result.Data.length !== 0 && !!result.Data[0].FileAs) {
                        const contact = result?.Data[0]
                        setContact(contact);
                    } else {
                        setFound(false)
                    }
                }
            );
        } catch (e) {
            console.error(e)
            setLoading(false)
        } 
    }

    return (
        <div>
            
            {(loading) ?
                <div className={css.loadingDiv}>
                    <ProgressIndicator label="Loading Agent Name" description="This tape will be destroyed after watching." />
                </div>
            : 
            <SearchBox
                styles={searchBoxStyles}
                placeholder="Search by email"
                onSearch={newValue => {onSearch(newValue)}}
          />
            }

            {!loading && contact?.FileAs &&
                <div style={{display: "flex", margin: 20}}>
                    <div style={{marginRight: 20}}>
                        <img src={`data:image/jpeg;base64,${contact.ProfilePicture}`} alt="profile picture" style={{width: contact.ProfilePictureWidth, height: contact.ProfilePictureHeight}} />
                    </div>
                    <div>
                        <table className="table">
                            <tbody>
                                <tr>
                                    <td>name:</td>
                                    <td>{contact.FirstName} {contact.LastName}</td>
                                </tr>
                                <tr>
                                    <td>business address:</td>
            <td>{contact.BusinessAddressStreet}, {contact.BusinessAddressPostalCode}, {contact.BusinessAddressCity}, {contact.BusinessAddressState}</td>
                                </tr>
                                <tr>
                                    <td>telephone number #1:</td>
                                    <td>{contact.TelephoneNumber1}</td>
                                </tr>
                                <tr>
                                    <td>webpage:</td>
                                    <td>{contact.WebPage}</td>
                                </tr>
                                <tr>
                                    <td>Skype:</td>
                                    <td>{contact.Skype}</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            }

            {!found && <div style={{marginLeft: 20, color: "#a44", fontWeight: 500}}>
                Name not found! 
                </div>
                }
        </div>
    );
}

export default App;

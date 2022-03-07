import * as React from "react";
import { useListItems, useCurrentUser } from "pnp-react-hooks";

export const ExampleComponent = () => {

    const currentUser = useCurrentUser();

    const items = useListItems<any>("Test", {
        query: {
            select: ["Title", "Id", "Author/Title"],
            expand: ["Author"],
            filter: `Author eq ${currentUser?.Id}`
        },
        disabled: !currentUser
    });

    return (<ul>
            { items?.map(item => (<li key={item.Id}>{item.Title}</li>)) }
            </ul>);
};
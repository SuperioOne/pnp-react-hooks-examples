import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType, Stack, Text } from "@fluentui/react";
import { useList } from "pnp-react-hooks";

interface ErrorDetails
{
    errorObj: Error;
    count: number;
}

export function ErrorHandling()
{
    const [error, setError] = React.useState<ErrorDetails>();
    const [refresh, setRefresh] = React.useState({});

    const errorHandler = (err: Error) =>
    {
        setError((previous) =>
        {
            return {
                count: (previous?.count ?? 0) + 1,
                errorObj: err,
            };
        });
    };

    useList("non existing list", {
        error: errorHandler,
    }, [refresh]);

    return (
        <Stack tokens={{ childrenGap: 25 }}>
            <Stack.Item>
                <Text variant="medium"> This example tries to get site with invalid URL configuration. </Text>
            </Stack.Item>
            <Stack.Item>
                <DefaultButton
                    text="Retry"
                    onClick={() => setRefresh({})}
                />
            </Stack.Item>
            <Stack.Item>
                {
                    error ?
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={true}
                        >
                            <Stack tokens={{ childrenGap: 3 }}>
                                <Stack.Item>
                                    {`Error Count : ${error.count}`}
                                </Stack.Item>
                                <Stack.Item>
                                    {error.errorObj.message}
                                </Stack.Item>
                            </Stack>
                        </MessageBar>
                        : undefined
                }
            </Stack.Item>
        </Stack>
    );
}
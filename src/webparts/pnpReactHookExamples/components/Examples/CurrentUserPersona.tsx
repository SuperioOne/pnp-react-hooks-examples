import * as React from "react";
import { DefaultButton, IPersonaSharedProps, Persona, PersonaSize, Shimmer, ShimmerElementsGroup, ShimmerElementType } from "@fluentui/react";
import { useCurrentUser, useProfile, useUser } from "pnp-react-hooks";

export function CurrentUserPersona()
{
    const [_, setState] = React.useState({});

    // get current user login name
    const curUser = useCurrentUser({
        query: {
            select: ["LoginName"]
        },
    });

    // disable hook until proper login name
    const profile = useProfile<IAzureADProfile>(curUser?.LoginName, {
        disabled: typeof curUser?.LoginName !== "string",
    });

    const testPro = useUser(profile?.Email, {
        disabled: typeof profile?.Email !== "string",
    });

    console.log(testPro);

    // Create persona information from profile
    const personaInfo: IPersonaSharedProps = React.useMemo(() =>
    {
        return profile
            ? {
                imageUrl: profile.PictureURL,
                showInitialsUntilImageLoads: true,
                text: profile.DisplayName,
                secondaryText: profile.Title,
                tertiaryText: profile.Email
            }
            : undefined;
    }, [profile]);

    return (
        <Shimmer
            customElementsGroup={getCustomElements()}
            width="150"
            isDataLoaded={personaInfo !== undefined}
        >
            <Persona {...personaInfo} size={PersonaSize.size40} />
            <DefaultButton text="test" onClick={() => setState({})} />
        </Shimmer>
    );
}

interface IAzureADProfile
{
    PictureURL: string | null;
    Title: string;
    DisplayName: string;
    Email: string;
}

const wrapperStyles = { display: 'flex' };

const getCustomElements = (): JSX.Element =>
{
    return (
        <div style={wrapperStyles}>
            <ShimmerElementsGroup
                shimmerElements={[
                    { type: ShimmerElementType.circle, height: 40 },
                    { type: ShimmerElementType.gap, width: 16, height: 40 },
                ]}
            />
            <ShimmerElementsGroup
                flexWrap
                width="100%"
                shimmerElements={[
                    { type: ShimmerElementType.line, width: '100%', height: 10, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.line, width: '90%', height: 8 },
                    { type: ShimmerElementType.gap, width: '10%', height: 20 },
                ]}
            />
        </div>
    );
};
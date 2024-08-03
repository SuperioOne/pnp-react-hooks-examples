/* eslint-disable @typescript-eslint/no-non-null-asserted-optional-chain */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import { IPersonaSharedProps, Persona, PersonaSize, Shimmer, ShimmerElementsGroup, ShimmerElementType } from "@fluentui/react";
import { useCurrentUser, useProfile } from "pnp-react-hooks";

interface IUserProfile {
    PictureURL: string | undefined;
    Title: string;
    DisplayName: string;
    Email: string;
}

const WRAPPER_STYLES = { display: 'flex' };

const getCustomElements = (): JSX.Element => {
    return (
        <div style={WRAPPER_STYLES}>
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
export default function CurrentUserPersona(): JSX.Element {
    // get current user login name
    const currentUser = useCurrentUser({
        query: {
            select: ["LoginName"]
        }
    });

    // load profile
    const profile = useProfile<IUserProfile>(currentUser?.LoginName!);

    // Create persona information from profile
    const personaInfo: IPersonaSharedProps | undefined = React.useMemo(() => {
        return profile
            ? {
                imageUrl: profile.PictureURL,
                showInitialsUntilImageLoads: true,
                text: profile.DisplayName,
                secondaryText: profile.Title,
                tertiaryText: profile.Email,
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
        </Shimmer>
    );
}



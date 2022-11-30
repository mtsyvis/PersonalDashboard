import { ServiceScope } from "@microsoft/sp-core-library";
import { createSlice, PayloadAction } from "@reduxjs/toolkit";
// import * as strings from "CodeAssistWebPartStrings";

// import CodesService from "../../../../services/CodesService";
import { IEmail } from "../../models/IEmail";
import { AppThunk, RootState } from "../store";

export interface IMainState {
    emailsLoaded: boolean;
    emails: IEmail[];
    error: string | undefined;
}

const initialState: IMainState = {
    emailsLoaded: false,
    emails: [],
    error: undefined
};

const mainSlice = createSlice({
    name: "mainSlice",
    initialState,
    reducers: {
        setEmailsLoaded: (state: IMainState, action: PayloadAction<boolean>) => {
            state.emailsLoaded = action.payload;
        },
        setEmails: (state: IMainState, action: PayloadAction<IEmail[]>) => {
            state.emails = action.payload;
        },
        setError: (state: IMainState, action: PayloadAction<string | undefined>) => {
            state.error = action.payload;
        }
    }
});

export const { setEmails, setEmailsLoaded, setError } = mainSlice.actions;

export const getCodes = (state: RootState): IEmail[] => state.main.emails;
export const getEmailsLoaded = (state: RootState): boolean => state.main.emailsLoaded;
export const getError = (state: RootState): string | undefined => state.main.error;


export default mainSlice.reducer;

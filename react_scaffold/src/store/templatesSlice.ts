import { createAsyncThunk, createSlice } from "@reduxjs/toolkit";
import type { PayloadAction } from "@reduxjs/toolkit";
import { templatesApi } from "../api/templatesApi";
import type { TemplateInfo } from "../types";

export type TemplatesState = {
  items: TemplateInfo[];
  status: "idle" | "loading" | "error";
};

const initialState: TemplatesState = {
  items: [],
  status: "idle"
};

export const fetchTemplates = createAsyncThunk("templates/fetch", async () => {
  const res = await templatesApi.list();
  return res;
});

const slice = createSlice({
  name: "templates",
  initialState,
  reducers: {
    setTemplates(state, action: PayloadAction<TemplateInfo[]>) {
      state.items = action.payload;
      state.status = "idle";
    }
  },
  extraReducers(builder) {
    builder
      .addCase(fetchTemplates.pending, (state) => {
        state.status = "loading";
      })
      .addCase(fetchTemplates.fulfilled, (state, action) => {
        state.items = action.payload;
        state.status = "idle";
      })
      .addCase(fetchTemplates.rejected, (state) => {
        state.status = "error";
      });
  }
});

export const { setTemplates } = slice.actions;
export const templatesReducer = slice.reducer;


import * as React from "react";
import styles from "./FeedBackForm.module.scss";
import type { IFeedBackFormProps } from "./IFeedBackFormProps";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import { Label } from "office-ui-fabric-react/lib/Label";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { css, Link } from "office-ui-fabric-react";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/fields/list";
import "@pnp/sp/items";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users/web";
import { IItemAddResult } from "@pnp/sp/items/types";

interface UserObject {
  Email: string;
  Expiration: string;
  Id: number;
  IsEmailAuthenticationGuestUser: boolean;
  IsHiddenInUI: boolean;
  IsShareByEmailGuestUser: boolean;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
}

interface IFeedBackFormStates {
  _typeOptions: IChoiceGroupOption[] | undefined;
  selectedType: string;
  Comments: string;
  CurrectUser?: UserObject;
  RequestSubmitted: boolean;
}

export default class FeedBackForm extends React.Component<
  IFeedBackFormProps,
  IFeedBackFormStates
> {
  /**
   *
   */
  constructor(props: IFeedBackFormProps) {
    super(props);
    this.state = {
      _typeOptions: [],
      selectedType: "",
      Comments: "",
      RequestSubmitted: false,
    };
  }
  GetChoiceFields = async () => {
    const list = this.props.sp.web.lists.getById(this.props.FeedbackListID);
    const r = await list.fields.getByInternalNameOrTitle(
      this.props.SupportType
    )();
    console.log("Fields", r.Choices);
    const Choices = r.Choices?.map((value) => ({ key: value, text: value }));
    this.setState({
      _typeOptions: Choices,
    });
  };

  public componentDidMount(): void {
    void this.GetChoiceFields();
    void this.getPerson();
  }

  public componentDidUpdate(
    prevProps: Readonly<IFeedBackFormProps>,
    prevState: Readonly<IFeedBackFormStates>,
    snapshot?: any
  ): void {
    if (this.props !== prevProps) {
      void this.GetChoiceFields();
    }
  }

  private _onConfigure = () => {
    // Context of the web part
    this.props.context.propertyPane.open();
  };

  private getPerson = async () => {
    // const user: UserObject = await this.props.sp.profiles.myProperties();
    const user: UserObject = await this.props.sp.web.currentUser();
    console.log("Profile fetched", user);
    this.setState({
      CurrectUser: user,
    });
  };

  private resetForm = () => {
    this.setState({
      Comments: "",
      selectedType: "",
      RequestSubmitted: false
    });
  }

  public render(): React.ReactElement<IFeedBackFormProps> {
    const { headingBackColor, displayMode, title, updateProperty } = this.props;

    return (
      <section className={css(styles.FeedBackForm, styles.msGrid)}>
        <div style={{backgroundColor: headingBackColor, padding:10, alignItems:'center'}}>
          <WebPartTitle
            className={styles.WebPartTitle}
            displayMode={displayMode}
            title={title}
            updateProperty={updateProperty}
          />
        </div>
        {/* {true ? ( */}
        {this.state.RequestSubmitted ? (
          <Placeholder
            iconName="SkypeCircleCheck"
            contentClassName={styles.SuccessPlaceHolder}
            iconText="Success!"
            description="Your Feedback has been Submitted"
            buttonLabel="Enter another?"
            hideButton={true}
          >
            {
              this.props.enableReEnterFormLink ? <div style={{textAlign:'center', padding: '0px 0px 10px 0px'}}>Want to submit another?<Link href="#" onClick={this.resetForm.bind(this)} > Click here.</Link></div> : <></>
            }
            </Placeholder>
        ) : !this.props.SupportType ? (
          <Placeholder
            iconName="Edit"
            iconText="Configure your web part"
            description="Please configure the web part."
            buttonLabel="Configure"
            onConfigure={this._onConfigure}
          />
        ) : (
          <>
            {/* Type selection */}
            <div className={css(styles.msGridrow)}>
              <Label className={css(styles.msGridcol25)} style={{textAlign:'right'}}>Type:</Label>
              {this.state._typeOptions &&
                this.state._typeOptions?.length > 1 && (
                  <ChoiceGroup
                    className={css(styles.msGridcol75)}
                    options={this.state._typeOptions}
                    onChange={this._onTypeChange}
                    required
                    styles={{ 
                      flexContainer: { 
                        display: "flex", 
                        flexDirection: "row",
                        justifyContent: 'space-between'
                      }, 
                      root: { 
                        width: 300,
                        height: 40
                      } 
                    }}                    
                  />
                )
              }
            </div>

            {/* Comment text area */}
            <div className={css(styles.msGridrow)} style={{alignItems: 'inherit'}}>
              <Label className={css(styles.msGridcol25)} style={{textAlign:'right'}}>Comment:</Label>
              <TextField
                className={css(styles.msGridcol50)}
                multiline
                rows={4}
                onChange={this._onCommentChange}
                required
              />
            </div>

            {/* Submit button */}
            <div className={css(styles.msGridrow)} style={{
              display: 'flex',
              alignItems:'center',
              flexDirection:'row',
              padding:'20px 0px'
            }}>
              <div className={styles.msGridcol75} style={{
                display: 'flex',
                justifyContent: 'flex-end'
              }}>
                <DefaultButton
                  className={css(styles.msGridcol25)}
                  text="Submit"
                  onClick={this._onSubmitClick}
                />
              </div>
            </div>
          </>
        )}
      </section>
    );
  }

  private _onTypeChange = (
    ev?: React.FormEvent<HTMLElement | HTMLInputElement>,
    option?: IChoiceGroupOption
  ): void => {
    // Handle type change
    // console.log("Type selected:", option?.key);
    this.setState({
      selectedType: option?.key ? option?.key : "",
    });
  };

  private _onCommentChange = (
    ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    // Handle comment change
    // console.log("Comment:", newValue);
    this.setState({
      Comments: newValue ? newValue : "",
    });
  };

  private _onSubmitClick = async (): Promise<void> => {
    // Handle submit
    // console.log("Submit clicked");
    const dynamicObject = {
      [this.props.colTitle]: this.state.CurrectUser?.Title,
      [this.props.colParticipant + "Id"]: this.state.CurrectUser?.Id,
      [this.props.colComments]: this.state.Comments,
      [this.props.colSupportType]: this.state.selectedType,
    };
    const iar: IItemAddResult = await this.props.sp.web.lists
      .getById(this.props.FeedbackListID)
      .items.add(dynamicObject);

    console.log(iar);
    if (iar) {
      this.setState({
        RequestSubmitted: true,
      });
    }
  };
}

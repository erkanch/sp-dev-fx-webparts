/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import {
	Panel,
	DialogFooter,
	DefaultButton,
	PrimaryButton,
} from 'office-ui-fabric-react';
import styles from './Styles.module.scss';

export interface ICustomPanelState {
	saving: boolean;
	placeName: string;
	folderPath: string;
	isOpen: boolean;
}

export interface ICustomPanelProps {
	onClose: () => void;
	onSave: (folderName: string) => void;
	onFolderClick: (folderName: string) => void;
	isOpen: boolean;
	currentTitle: string;
	items: any;
}

export default class CustomPanel extends React.Component<ICustomPanelProps, ICustomPanelState> {

	// private editedTitle: string = null;
	constructor(props: ICustomPanelProps) {
		super(props);
		this.state = {
			saving: false,
			placeName: "Site Pages",
			folderPath: "",
			isOpen: props.isOpen
		};
	}

	componentDidUpdate(prevProps: Readonly<ICustomPanelProps>, prevState: Readonly<ICustomPanelState>, snapshot?: any): void {
		if (prevProps.isOpen !== this.props.isOpen)
			this.setState({ isOpen: this.props.isOpen });
	}

	public render(): React.ReactElement<ICustomPanelProps> {
		const { isOpen } = this.state;
		const _hidePanel = (): void => {
			this.setState({ isOpen: false });
			this.props.onClose()
		};
		const _onSave = (folderRelativeUrl: string): void => {
			this.setState({ isOpen: false });
			this.props.onSave(folderRelativeUrl);
		}
		const _onFolderClick = (folderName: string, folderRelativeUrl: string): void => {
			this.props.onFolderClick(folderName);
			this.setState({
				placeName: folderName,
				folderPath: folderRelativeUrl
			})
		}
		return (
			<Panel closeButtonAriaLabel="Close" onDismiss={_hidePanel} isOpen={isOpen}>
				{this.state?.placeName ?
					<h2>{this.state.placeName}</h2>
				:
					<h2>Sub Folders</h2>
				}
				{this.props.items && this.props.items.length > 0 &&
					<div className={styles.subfoldersContainer}>
						<div className={styles.subfoldersHeader}>Choose a destination</div>
						{this.props.items.map((item) => {
							return (
								<div defaultValue={item?.Name} onClick={(e: any) => { return console.log(e) }} className={styles.subfolderWrapper}>
									<span>
										<img data-bind="attr:{src:iconUrl,alt:iconFieldAriaLabel},style:{width:size,height:size,'font-size':size}" src="https://res-1.cdn.office.net/files/fabric-cdn-prod_20220825.001/assets/item-types/16/folder.svg">
										</img>
										<DefaultButton href='' onClick={() => { _onFolderClick(item?.Name, item?.ServerRelativeUrl) }}>{item?.Name}</DefaultButton>
									</span>
								</div>
							)
						}
						)}
					</div>
				}
				<div>{this.state.folderPath}</div>
				<DialogFooter>
					<PrimaryButton text="Move here" onClick={() => { _onSave(this.state.folderPath) }} />
				</DialogFooter>
			</Panel>
		);
	}
}
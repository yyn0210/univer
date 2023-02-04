import { Tools } from '@univerjs/core';
import { Container, Content, Footer, Header, Layout, Sider } from '@univerjs/style-univer';
import cssVars from 'css-vars-ponyfill';
import defaultSkin from '@univerjs/style-univer/assets/css/skin/default.module.less';
import darkSkin from '@univerjs/style-univer/assets/css/skin/dark.module.less';
import greenSkin from '@univerjs/style-univer/assets/css/skin/green.module.less';
import { BaseComponentProps } from '../../BaseComponent';
import { AppContext } from '../../Common';
import { Component, createRef } from '../../Framework';
import style from './index.module.less';
import { ToolBar } from '../ToolBar';
import { ToolBar } from '../ToolBar';
import { InfoBar } from '../InfoBar';
import { UniverConfig } from '../../Basics/Interfaces/ComponentConfig/UniverConfig';
import { RightMenu } from '../RightMenu';
import { CountBar } from '../CountBar';
import { SheetBar } from '../SheetBar';
import { FormulaBar } from '../FormulaBar';
import { RichText } from '../RichText';

export interface BaseSheetContainerProps extends BaseComponentProps {
    config: UniverConfig;
    container?: HTMLElement;
    changeSkin: () => void;
    changeLocale: (locale: string) => void;
    mountCanvas?: (container: HTMLElement) => void;
    methods?: any;
}

// Types for state
interface IState {
    currentLocale: string;
    currentSkin: string;
}

/**
 * One universheet instance DOM container
 */
export class UniverContainer extends Component<BaseSheetContainerProps, IState> {
    leftContentLeft: number;

    leftContentTop: number;

    rightBorderX: number;

    rightBorderY: number;

    splitLeftRef = createRef<HTMLDivElement>();

    contentRef = createRef<HTMLDivElement>();

    constructor(props: BaseSheetContainerProps) {
        super(props, { context: props.config.context });
        const defaultSkin = props.config.skin ?? 'default';
        // init state
        this.state = {
            currentLocale: props.config.locale ?? 'zh',
            currentSkin: defaultSkin,
        };
        this.changeSkin(props.container!, defaultSkin);
    }

    componentDidMount() {
        this.props.getComponent?.(this);
        // this.props.mountCanvas?.(this.getContentRef().current!);
    }

    /**
     * split mouse down
     * @param e
     */
    handleSplitBarMouseDown = (e: MouseEvent) => {
        e = e || window.event; // Compatible with IE browser
        // Store the current mouse position
        this.leftContentLeft = this.splitLeftRef.current?.getBoundingClientRect().left!;
        this.leftContentTop = this.splitLeftRef.current?.getBoundingClientRect().top!;
        const mainContainer = this.splitLeftRef.current?.parentElement;
        this.rightBorderX = mainContainer?.getBoundingClientRect()?.width!;
        this.rightBorderY = mainContainer?.getBoundingClientRect()?.height!;
        document.addEventListener('mousemove', this.handleSplitBarMouseMove, false);
        document.addEventListener('mouseup', this.handleSplitBarMouseUp, false);
    };

    /**
     * split mouse move
     * @param e
     */
    handleSplitBarMouseMove = (e: MouseEvent) => {
        const layout = this.props.config.layout!;
        e = e || window.event; // Compatible with IE browser
        let diffLeft = e.clientX - this.leftContentLeft;
        let diffTop = e.clientY - this.leftContentTop;
        // Prevent crossing borders
        diffLeft = diffLeft >= this.rightBorderX ? this.rightBorderX : diffLeft;
        diffTop = diffTop >= this.rightBorderY ? this.rightBorderY : diffTop;
        // set new width
        if (layout.contentSplit === 'vertical') {
            this.splitLeftRef.current!.style.height = `${diffTop}px`;
        } else {
            this.splitLeftRef.current!.style.width = `${diffLeft}px`;
        }
    };

    /**
     * split mouse up
     * @param e
     */
    handleSplitBarMouseUp = (e: MouseEvent) => {
        document.removeEventListener('mousemove', this.handleSplitBarMouseMove, false);
        document.removeEventListener('mouseup', this.handleSplitBarMouseUp, false);
    };

    getContentRef() {
        return this.contentRef;
    }

    getSplitLeftRef() {
        return this.splitLeftRef;
    }

    handleChangeSkin(e: Event) {
        const { changeSkin } = this.props;
        const target = e.target as HTMLSelectElement;
        const skin = target.value;
        this.setState(
            {
                currentSkin: skin,
            },
            () => {
                this.changeSkin(this.props.container!, skin);
            }
        );
        changeSkin();
    }

    /**
     * Modify Dom Skin
     */
    changeSkin(container: HTMLElement, skin: string) {
        // Collect all  skins
        let root = document.documentElement;
        // get all skins
        const skins = {
            default: defaultSkin,
            dark: darkSkin,
            green: greenSkin,
        };
        // current skin set by user
        let currentSkin = skins[skin];
        // transform "primaryColor" to "--primary-color"
        currentSkin = Object.fromEntries(Object.keys(currentSkin).map((item) => [`--${item.replace(/([A-Z0-9])/g, '-$1').toLowerCase()}`, currentSkin[item]]));
        // ie11 does not support css variables, use css-vars-ponyfill to handle
        if (Tools.isIEBrowser()) {
            cssVars({
                // Options...
                // The container is invalid as rootElement, so the default setting is root.
                // Disadvantages: In ie11, only one set of skins can be used for multiple workbooks, and it is the skin set by the last workbook
                rootElement: root, // default
                variables: currentSkin,
            });
        } else {
            // set css variable
            let sheet = getSkinStyleSheet(container.id);
            /**
             *  covert object to style, remove " and replace , to ;
             *
             *  Example:
             *
             *  before: {--primary-color:"#0188fb",--primary-color-hover:"#5391ff"}
             *  after:  {--primary-color:#0188fb;--primary-color-hover:#5391ff;}
             */
            sheet.insertRule(
                `#${container.id} ${JSON.stringify(currentSkin)
                    .replace(/"/g, '')
                    .replace(/,(?=--)/g, ';')}`
            );
        }
        /**
         * get skin style sheet
         * @param id
         * @returns
         */
        function getSkinStyleSheet(id: string) {
            const title = 'universheet-skin-style';
            // avoid duplicates
            for (let i = 0; i < document.styleSheets.length; i++) {
                if (document.styleSheets[i].title === title) {
                    deleteStyleRuleIndexBySelector(document.styleSheets[i], id);
                    return document.styleSheets[i];
                }
            }
            const head = document.head || document.getElementsByTagName('head')[0];
            const styleEle = document.createElement('style');
            styleEle.title = title;
            head.appendChild(styleEle);
            return document.styleSheets[document.styleSheets.length - 1];
        }
        /**
         * delete style rule in universheet-skin-style
         * @param skinStyleSheet
         * @param id
         */
        function deleteStyleRuleIndexBySelector(skinStyleSheet: CSSStyleSheet, id: string) {
            let index = 0;
            for (let i = 0; i < skinStyleSheet.cssRules.length; i++) {
                const rule = skinStyleSheet.cssRules[i];
                if (rule instanceof CSSStyleRule && rule.selectorText === `#${id}`) {
                    index = i;
                    skinStyleSheet.deleteRule(index);
                    break;
                }
            }
        }
    }

    setLocale(e: Event) {
        const { changeLocale } = this.props;
        const target = e.target as HTMLSelectElement;
        const locale = target.value;
        // You must use setState to trigger the re-rendering of the child component
        this.setState(
            {
                currentLocale: locale,
            },
            () => {
                changeLocale(locale);
            }
        );
    }

    /**
     * Render the component's HTML
     *
     * @returns {void}
     */
    render() {
        const { methods } = this.props;
        const { context, layout } = this.props.config;
        const { currentLocale, currentSkin } = this.state;
        const config = layout!;
        // Set Provider for entire Container
        return (
            <AppContext.Provider
                value={{
                    context,
                }}
            >
                <Container className={style.layoutContainer}>
                    <Layout>
                        <Sider style={{ display: config.outerLeft ? 'block' : 'none' }}></Sider>
                        <Layout className={style.mainContent} style={{ position: 'relative' }}>
                            <Header style={{ display: config.header ? 'block' : 'none' }}>
                                {config.infoBar && <InfoBar {...methods.infoBar}></InfoBar>}
                                {config.toolBar && <ToolBar {...methods.toolbar}></ToolBar>}
                                {config.formulaBar && <FormulaBar {...methods.formulaBar}></FormulaBar>}
                            </Header>
                            <Layout>
                                <Sider
                                    style={{
                                        display: config.innerLeft ? 'block' : 'none',
                                    }}
                                >
                                    {/* innerLeft */}
                                </Sider>
                                <Content className={config.contentSplit === 'vertical' ? style.contentContainerVertical : style.contentContainerHorizontal}>
                                    {/* extend main content */}
                                    {/* <ModalGroup></ModalGroup> */}
                                    {!!config.contentSplit && (
                                        <Container ref={this.splitLeftRef} className={style.contentInnerLeftContainer}>
                                            <div className={style.hoverCursor} onMouseDown={this.handleSplitBarMouseDown}></div>
                                        </Container>
                                    )}
                                    <Container ref={this.contentRef} className={style.contentInnerRightContainer}>
                                        {config.rightMenu && <RightMenu {...methods.rightMenu}></RightMenu>}
                                        {config.cellEditor && <RichText {...methods.cellEditor}></RichText>}
                                        <div style={{ position: 'fixed', right: '200px', top: '10px', fontSize: '14px' }}>
                                            <span style={{ display: 'inline-block', width: 50, margin: '5px 0 0 5px' }}>皮肤</span>
                                            <select value={currentSkin} onChange={this.handleChangeSkin.bind(this)} style={{ width: 55 }}>
                                                <option value="default">默认</option>
                                                <option value="dark">暗黑</option>
                                                <option value="green">绿色</option>
                                            </select>
                                            <span style={{ display: 'inline-block', width: 50, margin: '5px 0 0 5px' }}>语言</span>
                                            <select value={currentLocale} onChange={this.setLocale.bind(this)} style={{ width: 55 }}>
                                                <option value="en">English</option>
                                                <option value="zh">中文</option>
                                            </select>
                                        </div>
                                    </Container>
                                </Content>
                                <Sider
                                    style={{
                                        display: config.innerRight ? 'block' : 'none',
                                    }}
                                >
                                    {/* innerRight */}
                                    {/* <SideGroup></SideGroup> */}
                                </Sider>
                            </Layout>
                            <Footer
                                style={{
                                    display: config.footer ? 'block' : 'none',
                                }}
                            >
                                {config.sheetBar && <SheetBar {...methods.sheetBar}></SheetBar>}
                                {config.countBar && <CountBar {...methods.countBar}></CountBar>}
                            </Footer>
                        </Layout>
                        <Sider
                            style={{
                                display: config.outerRight ? 'block' : 'none',
                            }}
                            className={style.outerRightContainer}
                        ></Sider>
                    </Layout>
                </Container>
            </AppContext.Provider>
        );
    }
}

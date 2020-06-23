import * as React from "react";
import "./FlyOut.sass";
import { Session } from "@pnp/sp-taxonomy";

class FlyOut extends React.Component<any, any>{
    constructor(props) {
        super(props);
    }

    async componentDidMount() {
        const taxonomy = new Session("https://technovert2020.sharepoint.com/sites/Technovert");
        let termStoresList = await taxonomy.termStores.get();
        let termsMenu = [];
        for (var i = 0; i < termStoresList.length; i++) {
            let termStoreName = termStoresList[i].Name;
            let termGroups = termStoresList[i].getTermGroupById("")
            console.log("Term Group : ", termGroups);
        }
        console.log(termStoresList);
    }

    render() {
        return (
            <div className="row">
                <ul className="dropdown-menu multi-level" role="menu">
                    <li className="dropdown-submenu">
                        <a data-tabindex="-1" href="#">Home</a>
                        <ul className="dropdown-menu">
                            <li><a href="#">Second level</a></li>
                            <li><a href="#">Second level</a></li>
                            <li><a href="#">Second level</a></li>
                        </ul>
                    </li>
                    <li className="dropdown-submenu">
                        <a data-tabindex="-1" href="#">Footer</a>
                        <ul className="dropdown-menu">
                            <li><a data-tabindex="-1" href="#">Second level</a></li>
                            <li className="dropdown-submenu">
                                <a href="#">Even More..</a>
                                <ul className="dropdown-menu">
                                    <li><a href="#">3rd level</a></li>
                                    <li><a href="#">3rd level</a></li>
                                </ul>
                            </li>
                            <li><a href="#">Second level</a></li>
                            <li><a href="#">Second level</a></li>
                        </ul>
                    </li>
                    <li className="dropdown-submenu">
                        <a data-tabindex="-1" href="#">Not Assigned</a>
                        <ul className="dropdown-menu">
                            <li><a data-tabindex="-1" href="#">Second level</a></li>
                            <li className="dropdown-submenu">
                                <a href="#">Even More..</a>
                                <ul className="dropdown-menu">
                                    <li><a href="#">3rd level</a></li>
                                    <li><a href="#">3rd level</a></li>
                                </ul>
                            </li>
                            <li><a href="#">Second level</a></li>
                            <li><a href="#">Second level</a></li>
                        </ul>
                    </li>
                </ul>
            </div>
        )
    }
}

export default FlyOut;
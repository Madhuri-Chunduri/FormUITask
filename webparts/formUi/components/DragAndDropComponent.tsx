import * as React from 'react';

class DragAndDrop extends React.Component<any, any> {
    private readonly dropRef: React.RefObject<HTMLInputElement>

    constructor(props) {
        super(props);
        this.state = {
            drag: true
        }
        this.dropRef = React.createRef();
    }

    handleDrop = (e) => {
        e.preventDefault()
        e.stopPropagation()
        this.setState({ drag: false })
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            this.props.handleDrop(e.dataTransfer.files)
            e.dataTransfer.clearData()
            this.dragCounter = 0
        }
    }

    dragCounter: any
    componentDidMount() {
        let div = this.dropRef.current
        div.addEventListener('drop', this.handleDrop)
    }

    componentWillUnmount() {
        let div = this.dropRef.current
        div.removeEventListener('drop', this.handleDrop)
    }

    render() {
        return (
            <div>
                <div ref={this.dropRef} className="dragArea">
                    {this.props.children}
                </div>
            </div>
        )
    }
}
export default DragAndDrop
# `uiVBA`

A [`stdVBA`](https://github.com/sancarn/stdVBA) project to build a [`react`](https://react.dev/)-like set of UI components for VBA.

## Vision

```vb
Class Video
  Private Type TThis
    url as string
    title as string
    description as string
  End Type
  Private This as TThis

  'Constructors
  Public Function Create(ByVal sTitle as string, ByVal sDescription as string, ByVal sURL as string) as Video
    Set Create = new Video
    Call Create.protInit(sTitle, sDescription, sURL)
  End Function
  Public Sub protInit(ByVal sTitle as string, ByVal sDescription as string, ByVal sURL as string)
    This.url = sURL
    This.title = sTitle
    This.description = sDescription
  End Sub

  'Component
  Public Function Render(ui)
    With ui
      With .Add(uiDiv.Create())
        Call .Add(uiThumbnail.Create(this.url))
        With .Add(uiLink.Create(this.url))
          Call .Add(uiTitle.Create(content := this.title))
          Call .Add(uiParagraph.Create(content := this.description))
        End With
        Call .Add(uiLikeButton.Create(this.url))
      End With
    End With
  End Function
End Class
```

Analagous to

```js
function Video({ video }) {
  return (
    <div>
      <Thumbnail video={video} />
      <a href={video.url}>
        <h3>{video.title}</h3>
        <p>{video.description}</p>
      </a>
      <LikeButton video={video} />
    </div>
  );
}
```


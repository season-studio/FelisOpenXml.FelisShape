using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.IO;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The class of the presentation
    /// </summary>
    public class FelisPresentation : IDisposable
    {
        /// <summary>
        /// The document of the presentation
        /// </summary>
        public readonly PresentationDocument SDKDocument;
        /// <summary>
        /// The stream assigned to the document
        /// </summary>
        public readonly Stream DocumentStream;
        /// <summary>
        /// The part of the presentation
        /// </summary>
        public readonly PresentationPart SDKPart;
        /// <summary>
        /// The root element of the presentation
        /// </summary>
        public readonly P.Presentation SDKRootElement;
        /// <summary>
        /// Record the marster part copied from any other presentation
        /// </summary>
        protected readonly ConditionalWeakTable<P.Presentation, Dictionary<string, string>> masterPartCopiedRecord = new ConditionalWeakTable<P.Presentation, Dictionary<string, string>>();
        private bool disposed;

        /// <summary>
        /// Create a presentation by loading the data from the given stream
        /// </summary>
        /// <param name="_sourceStream">The stream containing the source data.</param>
        public FelisPresentation(Stream _sourceStream)
        {
            disposed = false;

            DocumentStream = _sourceStream;
            if (_sourceStream.CanSeek)
            {
                _sourceStream.Seek(0, SeekOrigin.Begin);
            }
            SDKDocument = PresentationDocument.Open(_sourceStream, _sourceStream.CanWrite);
            var part = SDKDocument.PresentationPart;
            var element = part?.Presentation;
            if (null == element)
            {
                throw new OpenXmlPackageException($"Part({null != part}):Root({null != element})");
            }
            SDKPart = part!;
            SDKRootElement = element;
        }

        /// <summary>
        /// Create an empty presentation
        /// </summary>
        /// <param name="_fnCustomInit">The action for initializing the empty presentation. The default action will be taken when this argument is null.</param>
        public FelisPresentation(Action<FelisPresentation>? _fnCustomInit = null)
        {
            disposed = false;

            DocumentStream = new MemoryStream();
            SDKDocument = PresentationDocument.Create(DocumentStream, PresentationDocumentType.Presentation);
            SDKPart = SDKDocument.AddPresentationPart();
            SDKRootElement = (SDKPart.Presentation = new P.Presentation(
                new P.SlideMasterIdList(),
                new P.SlideIdList()
            ));

            var cpPart = SDKDocument.AddCoreFilePropertiesPart();
            using (var cpPartWriter = new OpenXmlPartWriter(cpPart))
            {
                cpPartWriter.WriteStartDocument(true);
                cpPartWriter.WriteElement(OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" />"));
            }

            _fnCustomInit?.Invoke(this);

            OpenXmlElement?[] complements = new OpenXmlElement?[3];
            int count = 0;

            if (null == SDKRootElement.GetFirstChild<P.SlideSize>())
            {
                complements[count++] = new P.SlideSize()
                {
                    Cx = 12192000,
                    Cy = 6858000
                };
            }

            if (null == SDKRootElement.GetFirstChild<P.NotesSize>())
            {
                complements[count++] = new P.NotesSize()
                {
                    Cx = 6858000,
                    Cy = 9144000
                };
            }

            if (null == SDKRootElement.GetFirstChild<P.DefaultTextStyle>())
            {
                complements[count++] = new P.DefaultTextStyle();
            }

            if (count > 0)
            {
                SDKRootElement.Append(complements.Take(count)!);
            }

            SDKRootElement.Save();
        }

        /// <summary>
        /// Finalize
        /// </summary>
        ~FelisPresentation()
        {
            Dispose(false);
        }

        /// <summary>
        /// Initialize a presentation from an other presentation.
        /// This function can be used when constructing the presentation.
        /// </summary>
        /// <param name="_target"></param>
        /// <param name="_template"></param>
        public static void InitializePresentaionFromOther(FelisPresentation _target, FelisPresentation _template)
        {
            FSUtilities.CopyPart(_template.SDKDocument.ExtendedFilePropertiesPart, () => _target.SDKDocument.AddExtendedFilePropertiesPart());

            FSUtilities.CopyPart(_template.SDKPart.PresentationPropertiesPart, () => _target.SDKPart.AddNewPart<PresentationPropertiesPart>());
            FSUtilities.CopyPart(_template.SDKPart.ViewPropertiesPart, () => _target.SDKPart.AddNewPart<ViewPropertiesPart>());
            FSUtilities.CopyPart(_template.SDKPart.TableStylesPart, () => _target.SDKPart.AddNewPart<TableStylesPart>());

            OpenXmlElement?[] addings = new OpenXmlElement?[3];
            int count = 0;

            var slideSize = _template.SDKRootElement.SlideSize?.CloneNode(true) as P.SlideSize;
            if (null != slideSize)
            {
                addings[count++] = slideSize;
            }
            var notesSize = _template.SDKRootElement.NotesSize?.CloneNode(true) as P.NotesSize;
            if (null != notesSize)
            {
                addings[count++] = notesSize;
            }
            var defTextStyle = _template.SDKRootElement.DefaultTextStyle?.CloneNode(true) as P.DefaultTextStyle;
            if (null != defTextStyle)
            {
                addings[count++] = defTextStyle;
            }

            if (count > 0)
            {
                _target.SDKRootElement.Append(addings.Take(count)!);
            }
        }

        /// <summary>
        /// Create a empty presentation taking an existed presentaion as the template
        /// </summary>
        /// <param name="_template"></param>
        /// <returns></returns>
        public static FelisPresentation From(FelisPresentation _template)
        {
            var ret = new FelisPresentation((target) =>
            {
                InitializePresentaionFromOther(target, _template);
            });

            return ret;
        }

        /// <summary>
        /// Save the presentation
        /// </summary>
        /// <param name="_stream">The destination stream for saving the presentation. The origin stream of the document will be taken as the default if this argument is null.</param>
        public void Save(Stream? _stream = null)
        {
            SDKDocument.Save();
            if ((null != _stream) && (_stream != DocumentStream))
            {
                SDKDocument.Clone(_stream);
            }
        }

        /// <summary>
        /// Save the presentation to a file with the special file name
        /// </summary>
        /// <param name="_filePath"></param>
        public void Save(string _filePath)
        {
            using (var f = File.Open(_filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                Save(f);
            }
        }

        /// <summary>
        /// Close the presentation
        /// </summary>
        public void Close()
        {
            Dispose();
        }

        /// <summary>
        /// Disposing the resource
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposing the resource
        /// </summary>
        /// <param name="_isDispose"></param>
        protected virtual void Dispose(bool _isDispose)
        {
            if (_isDispose && !disposed)
            {
                SDKDocument.Close();
                DocumentStream.Dispose();
                disposed = true;
            }
        }

        /// <summary>
        /// Get an enumerator for all the slides contained in this presentation
        /// </summary>
        public IEnumerable<FelisSlide> Slides
        {
            get
            {
                var slideIds = SDKRootElement.SlideIdList?.ChildElements;
                if (null != slideIds)
                {
                    foreach (var idElement in slideIds)
                    {
                        if ((idElement is P.SlideId id) && (!string.IsNullOrWhiteSpace(id.RelationshipId)))
                        {
                            var relPart = SDKPart.GetPartById(id.RelationshipId!);
                            if (relPart is SlidePart slidePart)
                            {
                                var slide = FelisSlide.From(this, slidePart);
                                if (null != slide)
                                {
                                    yield return slide;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Insert a existed slide into the presentation
        /// </summary>
        /// <param name="_sourceSlide"></param>
        /// <param name="_index"></param>
        /// <returns></returns>
        public FelisSlide? InsertSlide(FelisSlide _sourceSlide, int _index = -1)
        {
            lock (SDKDocument)
            {
                FelisPresentation sourcePresObj = _sourceSlide.Presentation;
                PresentationDocument sourceDoc = sourcePresObj.SDKDocument;
                PresentationPart? sourcePresPart = sourcePresObj.SDKPart;
                P.Presentation sourcePres = sourcePresObj.SDKRootElement;
                SlidePart sourceSlidePart = _sourceSlide.SDKPart;
                SlidePart? destSlidePart;
                bool isInSameDoc = (sourcePresObj == this);

                if (isInSameDoc)
                {
                    // 如果Slide来自同一文档，则把Slide先复制到一个临时的内存文档中，再复制回来
                    // 这样做的原因是：
                    // Slide中可能含有需要独立存在的关联部件，比如ChartPart，这些部件需要逐一复制，比较繁琐，
                    // 而且Open XML SDK中似乎存在BUG，比如找不到可以在/ppt/charts下创建ChartPart部件的API，对手动逐一复制关联部件造成了阻碍
                    using (var tmpPres = new FelisPresentation())
                    {
                        var tmpSlidePart = tmpPres.SDKPart.AddPart(sourceSlidePart);
                        destSlidePart = SDKPart.AddPart(tmpSlidePart);
                    }
                }
                else
                {
                    destSlidePart = SDKPart.AddPart(sourceSlidePart);
                }

                if (null != destSlidePart)
                {
                    // 确保Presentation中的关键清单元素存在
                    if (null == SDKRootElement.SlideIdList)
                    {
                        SDKRootElement.SlideIdList = new P.SlideIdList();
                    }
                    P.SlideIdList slideIdList = SDKRootElement.SlideIdList!;
                    if (null == SDKRootElement.SlideMasterIdList)
                    {
                        SDKRootElement.SlideMasterIdList = new P.SlideMasterIdList();
                    }

                    // 确定新的Slide列表中的ID编号
                    uint id = 256;
                    foreach (var scanItem in slideIdList)
                    {
                        if (scanItem is P.SlideId scanId)
                        {
                            if (scanId.Id?.Value > id)
                            {
                                id = scanId.Id;
                            }
                        }
                    }

                    // 插入新Slide的引用到胶片清单中
                    int destIdx = (_index >= 0) ? Math.Min(_index, slideIdList.ChildElements.Count) : Math.Max(slideIdList.ChildElements.Count + 1 + _index, 0);
                    slideIdList.InsertAt(new P.SlideId()
                    {
                        Id = id + 1,
                        RelationshipId = SDKPart.GetIdOfPart(destSlidePart)
                    }, destIdx);

                    if (isInSameDoc)
                    {
                        // 源Slide来自本文档，则直接重设母版关联关系
                        destSlidePart.DeleteParts(destSlidePart.GetPartsOfType<SlideLayoutPart>());
                        if (null != sourceSlidePart.SlideLayoutPart)
                        {
                            destSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
                        }
                        var destNotePart = destSlidePart.NotesSlidePart;
                        if (null != destNotePart)
                        {
                            destNotePart.DeleteParts(destNotePart.GetPartsOfType<NotesMasterPart>());
                            if (null != sourceSlidePart.NotesSlidePart?.NotesMasterPart)
                            {
                                destNotePart!.CreateRelationshipToPart(sourceSlidePart.NotesSlidePart!.NotesMasterPart!);
                            }
                        }
                    }
                    else
                    {
                        // 源Slide来自其他Presentation，则修正关联母版的插入关系
                        var masterCopiedList = masterPartCopiedRecord!.GetOrCreateValue(sourcePres);
                        lock (masterCopiedList)
                        {
                            // 先处理Slide母版
                            var sourceMasterPart = sourceSlidePart.SlideLayoutPart?.SlideMasterPart;
                            if (null != sourceMasterPart)
                            {
                                // 定位当前Slide关联的母版
                                string oriMasterPartUri = sourceMasterPart.Uri.ToString();
                                int layoutIndex = 0;
                                foreach (var layoutPart in sourceMasterPart.SlideLayoutParts)
                                {
                                    if (layoutPart.Uri == sourceSlidePart.SlideLayoutPart!.Uri)
                                    {
                                        break;
                                    }
                                    layoutIndex++;
                                }

                                // 查找母版是否已经被插入过
                                if (masterCopiedList.TryGetValue(oriMasterPartUri, out string? masterPartId))
                                {
                                    // 母版已被添加，则用已添加的母版工作
                                    SlideMasterPart? destMasterPart = SDKPart.GetPartById(masterPartId!) as SlideMasterPart;
                                    SlideLayoutPart? destLayoutPart = destMasterPart?.SlideLayoutParts?.ElementAtOrDefault(layoutIndex);
                                    if (null != destLayoutPart)
                                    {
                                        destSlidePart.DeleteParts(destSlidePart.GetPartsOfType<SlideLayoutPart>());
                                        destSlidePart.AddPart(destLayoutPart);
                                    }
                                }
                                else
                                {
                                    // 母版未被添加，则添加母版
                                    SlideMasterPart destMasterPart = SDKPart.AddPart(destSlidePart.SlideLayoutPart!.SlideMasterPart!);
                                    id = int.MaxValue;
                                    foreach (P.SlideMasterId scanId in SDKRootElement.SlideMasterIdList!)
                                    {
                                        if (scanId.Id?.Value > id)
                                        {
                                            id = scanId.Id;
                                        }
                                    }
                                    P.SlideMasterId slideMaterId = new P.SlideMasterId()
                                    {
                                        Id = ++id,
                                        RelationshipId = SDKPart.GetIdOfPart(destMasterPart)
                                    };
                                    SDKRootElement.SlideMasterIdList!.Append(slideMaterId);
                                    foreach (SlideMasterPart slideMasterPart in SDKPart.SlideMasterParts)
                                    {
                                        foreach (P.SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList!)
                                        {
                                            slideLayoutId.Id = ++id;
                                        }

                                        slideMasterPart.SlideMaster.Save();
                                    }

                                    masterCopiedList.Add(oriMasterPartUri, slideMaterId.RelationshipId!);
                                }
                            }

                            // 再处理备注母版
                            var sourceNodeMasterPart = sourceSlidePart.NotesSlidePart?.NotesMasterPart;
                            if ((null != sourceNodeMasterPart) && (destSlidePart.NotesSlidePart is NotesSlidePart destNotePart))
                            {
                                var oriPartUri = sourceNodeMasterPart.Uri.ToString();
                                if (masterCopiedList.TryGetValue(oriPartUri, out string? partDestUri))
                                {
                                    // 母版已被添加过，则延用
                                    var destPart = SDKPart.GetPartsOfType<SlidePart>().FirstOrDefault(e => (e != destSlidePart) && (e.NotesSlidePart?.NotesMasterPart?.Uri.ToString() == partDestUri))?.NotesSlidePart?.NotesMasterPart;
                                    if (null != destPart)
                                    {
                                        destNotePart.DeleteParts(destNotePart.GetPartsOfType<NotesMasterPart>());
                                        destNotePart.AddPart(destPart);
                                    }
                                }
                                else
                                {
                                    // 母版未被添加过，则登记
                                    masterCopiedList.Add(oriPartUri, destSlidePart.NotesSlidePart?.NotesMasterPart?.Uri.ToString()??string.Empty);
                                }
                            }
                        }

                        // 源Slide来自其他Presentation，还要修复关系部件对Presentation的关系
                        var sourceRefParts = sourceSlidePart.Parts.ToArray();
                        if (null != sourceRefParts)
                        {
                            foreach (var refInfo in destSlidePart.Parts)
                            {
                                var oriPart = sourceRefParts.SingleOrDefault(e => e.RelationshipId == refInfo.RelationshipId)?.OpenXmlPart;
                                if ((null != oriPart)
                                    && (oriPart.ContentType == refInfo.OpenXmlPart.ContentType)
                                    && (oriPart.RelationshipType == refInfo.OpenXmlPart.RelationshipType)
                                    && oriPart.GetParentParts().Where(e => e is PresentationPart).Any())
                                {
                                    SDKPart.CreateRelationshipToPart(refInfo.OpenXmlPart);
                                }
                            }
                        }
                    }

                    // 保存
                    SDKRootElement.Save();

                    return FelisSlide.From(this, destSlidePart);
                }

                return null;
            }
        }

        /// <summary>
        /// Delete an existed slide in the presentation
        /// </summary>
        /// <param name="_slide"></param>
        public void RemoveSlide(FelisSlide _slide)
        {
            lock (SDKDocument)
            {
                if (_slide.Presentation != this)
                {
                    return;
                }

                if (null == SDKRootElement.SlideIdList)
                {
                    SDKRootElement.Append(new P.SlideIdList());
                }
                P.SlideIdList slideIdList = SDKRootElement.SlideIdList!;

                P.SlideId? slideId = null;
                foreach (var scanItem in slideIdList)
                {
                    if ((scanItem is P.SlideId scanId) && !string.IsNullOrEmpty(scanId.RelationshipId))
                    {
                        var part = SDKPart.GetPartById(scanId.RelationshipId!);
                        if ((null != part) && (part.Uri == _slide.SDKPart.Uri))
                        {
                            slideId = scanId;
                            break;
                        }
                    }
                }

                if (null != slideId)
                {
                    slideId.Remove();

                    if (SDKRootElement.CustomShowList != null)
                    {
                        foreach (var customShow in SDKRootElement.CustomShowList.Elements<CustomShow>())
                        {
                            if (customShow.SlideList != null)
                            {
                                LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                                foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                                {
                                    if (slideListEntry.Id != null && slideListEntry.Id == slideId.RelationshipId)
                                    {
                                        slideListEntries.AddLast(slideListEntry);
                                    }
                                }

                                foreach (SlideListEntry slideListEntry in slideListEntries)
                                {
                                    customShow.SlideList.RemoveChild(slideListEntry);
                                }
                            }
                        }
                    }

                    SDKRootElement.Save();

                    SDKPart.DeletePart(_slide.SDKPart);
                }
            }
        }

        /// <summary>
        /// Check if the presentaion is writable
        /// </summary>
        public bool Writable => ((SDKPart.OpenXmlPackage.Package.FileOpenAccess & FileAccess.Write) != 0);
    }
}

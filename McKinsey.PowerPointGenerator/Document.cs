using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using McKinsey.PowerPointGenerator.Extensions;
using NLog;

namespace McKinsey.PowerPointGenerator
{
    public class Document : IDisposable
    {
        private List<SlideElement> slides;
        internal PresentationDocument PptDocument { get; private set; }
        internal bool IsLoaded { get; set; }

        /// <summary>
        /// Gets the PresentationPart object.
        /// </summary>
        /// <value>
        /// The PresentationPart.
        /// </value>
        internal PresentationPart PresentationPart
        {
            get
            {
                if (PptDocument != null)
                {
                    return PptDocument.PresentationPart;
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the listof slides.
        /// </summary>
        /// <value>
        /// The slides.
        /// </value>
        public virtual List<SlideElement> Slides
        {
            get
            {
                if (PresentationPart != null)
                {
                    if (slides != null)
                    {
                        return slides;
                    }
                }
                return null;
            }
        }

        ~Document()
        {
            Dispose(false);
        }

        /// <summary>
        /// Loads document from the stream
        /// </summary>
        /// <param name="stream">The stream.</param>
        public void Load(Stream stream)
        {
            if (!IsLoaded)
            {
                PptDocument = PresentationDocument.Open(stream, true, new OpenSettings { AutoSave = true });
                IsLoaded = true;
            }
        }

        public void SaveAndClose()
        {
            this.Dispose();
        }

        /// <summary>
        /// Gets all slides. Hidden slides are ignored.
        /// </summary>
        /// <returns></returns>
        internal virtual List<SlideElement> GetSlides()
        {
            Logger logger = LogManager.GetLogger("Generator");
            slides = new List<SlideElement>();
            var slideIds = PresentationPart.Presentation.SlideIdList.Elements<SlideId>();
            int slideNumber = 0;
            foreach (SlidePart slidePart in slideIds.Select(s => GetSlide(s)))
            {
                if (!slidePart.IsHidden())
                {
                    slides.Add(new SlideElement(this) { Slide = slidePart.Slide, Number = slideNumber });
                }
                else
                {
                    logger.Debug("Skipped slide {0} because it's hidden", slideNumber);
                }
                slideNumber++;
            }
            return slides;
        }

        /// <summary>
        /// Gets the slide from the SlideId using RelationshipId
        /// </summary>
        /// <param name="slideId">The slide identifier.</param>
        /// <returns></returns>
        internal SlidePart GetSlide(SlideId slideId)
        {
            ParameterHelpers.NotNull<SlideId>(slideId, "slideId");

            if (PresentationPart != null)
            {
                return (SlidePart)PresentationPart.GetPartById(slideId.RelationshipId);
            }
            return null;
        }

        private bool disposed;

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    if (PptDocument != null)
                    {
                        PresentationPart.Presentation.Save();
                        PptDocument.Close();
                        IsLoaded = false;
                    }
                }
                disposed = true;
            }
        }
    }
}
